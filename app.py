# app.py

import pandas as pd
import re
from flask import Flask, render_template, request, redirect, url_for
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.exc import IntegrityError

# --- Configuration & Initialization ---
app = Flask(__name__)
XLSX_FILE_NAME = 'Fresh_10.12.2025.xlsx'

# Database Configuration (SQLite)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///inventory.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)


# Add a thousands separator filter to Jinja2
def thousands_separator(value):
    return f"{value:,}"


app.jinja_env.filters['thousands_separator'] = thousands_separator


# --- Database Model Definition ---
class Asset(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    asset_type = db.Column('Asset Type', db.String(100), nullable=False)
    product = db.Column('Product', db.String(255))
    name = db.Column('Name', db.String(255))
    serial_number = db.Column('Serial Number', db.String(255), unique=True, nullable=True, index=True)
    used_by_name = db.Column('Used by (Name)', db.String(255))
    used_by_email = db.Column('Used by (Email)', db.String(255), index=True)
    department = db.Column('Department', db.String(255))
    location = db.Column('Location', db.String(255))


# --- Utility: Helper function for cleaning column headers ---
def clean_col_name(col):
    """Standardizes column names for consistent matching (e.g., 'Serial Number' -> 'serial_number')."""
    if isinstance(col, str):
        # Remove non-word/space chars, replace spaces with underscores, convert to lower case
        cleaned = re.sub(r'[^\w\s]', '', col).strip().replace(' ', '_')
        return cleaned.lower()
    return str(col).lower()


# --- Data Loading and Database Setup Function ---

REVERSE_MAPPING = {
    'product': ['product', 'model'],
    'name': ['name', 'description'],
    'serial_number': ['serial_number', 'serial_num', 'serial'],
    'used_by_name': ['used_by_name'],
    'used_by_email': ['used_by', 'used_by_email'],
    'department': ['department', 'dept'],
    'location': ['location', 'loc']
}
LOWERCASE_FIELDS = ['serial_number', 'asset_type', 'department', 'location', 'used_by_email']


def setup_database_from_excel():
    """
    Initializes the database and populates it with data from the XLSX file,
    using dynamic column mapping and committing per sheet for resilience.
    """
    with app.app_context():
        db.create_all()

        if Asset.query.count() == 0:
            print(f"Database is empty. Populating from '{XLSX_FILE_NAME}'...")
            try:
                # CRUCIAL: Read all data as strings to prevent '0' issue
                all_sheets = pd.read_excel(
                    XLSX_FILE_NAME,
                    sheet_name=None,
                    engine='openpyxl',
                    dtype=str
                )

                records_added_total = 0
                existing_serials = set(
                    [s[0] for s in db.session.query(Asset.serial_number).filter(Asset.serial_number.isnot(None)).all()]
                )

                for sheet_name, df in all_sheets.items():

                    if sheet_name.lower().startswith('unnamed:') or df.empty:
                        continue

                        # 1. Clean and standardize all column names
                    df.columns = [clean_col_name(col) for col in df.columns]

                    # 2. Drop irrelevant 'unnamed:' columns
                    unnamed_cols = [col for col in df.columns if col.startswith('unnamed')]
                    if unnamed_cols:
                        df = df.drop(columns=unnamed_cols)

                    # 3. Build a dynamic column map for this specific sheet
                    sheet_map = {}
                    for model_attr, possible_cols in REVERSE_MAPPING.items():
                        for col in df.columns:
                            if col in possible_cols:
                                # Map the actual cleaned column name to the model attribute name
                                sheet_map[col] = model_attr
                                break

                    sheet_records_added = 0
                    sheet_asset_type = sheet_name.lower().strip()

                    try:
                        # 4. Process rows for the current sheet
                        for index, row in df.iterrows():

                            # CRUCIAL FIX: Try-except block for individual rows
                            try:
                                asset_data = {'asset_type': sheet_asset_type}

                                # Extract and clean data using the sheet-specific map
                                for col_name, model_attr in sheet_map.items():
                                    try:
                                        value = str(row[col_name]).strip()
                                    except KeyError:
                                        continue
                                    except Exception:
                                        value = None

                                    # Handle empty/null values
                                    if value is None or value.lower() in ('nan', '', 'none', 'nan '):
                                        value = None

                                    # Enforce lowercase on key fields
                                    if model_attr in LOWERCASE_FIELDS and value is not None:
                                        value = value.lower()

                                    asset_data[model_attr] = value

                                serial_num = asset_data.get('serial_number')

                                # FIX: Generate synthetic serial number if missing (for consumables)
                                if serial_num is None:
                                    sheet_prefix = sheet_name.replace(' ', '_').upper()
                                    serial_num = f"SYNTHETIC_{sheet_prefix}_{index + 1}"
                                    asset_data['serial_number'] = serial_num

                                # Check if the serial (real or synthetic) is already known
                                if serial_num in existing_serials:
                                    continue  # Skip if already seen

                                # Add to session
                                new_asset = Asset(**asset_data)
                                db.session.add(new_asset)
                                existing_serials.add(serial_num)
                                sheet_records_added += 1

                            except Exception as row_e:
                                # This row failed to process, log it and continue to the next row
                                print(
                                    f"Warning: Skipping row {index + 2} in sheet '{sheet_name}' due to error: {row_e}")

                        # 5. COMMIT THE CURRENT SHEET'S DATA
                        db.session.commit()
                        records_added_total += sheet_records_added
                        print(f"Successfully loaded {sheet_records_added} assets from sheet: '{sheet_name}'.")

                    # Catch sheet-level database errors (should only be IntegrityError now)
                    except IntegrityError as e:
                        db.session.rollback()
                        print(
                            f"FAILED to load sheet '{sheet_name}' due to IntegrityError (duplicate serials). Rolling back sheet changes. Error: {e}")

                    # Catch any remaining unexpected error during sheet processing
                    except Exception as e:
                        db.session.rollback()
                        print(
                            f"FAILED to load sheet '{sheet_name}' due to UNEXPECTED error. Rolling back sheet changes. Error: {e}")

                print(f"Total unique assets added to the database: {records_added_total}.")

            except FileNotFoundError:
                print(f"Error: Initial Excel file '{XLSX_FILE_NAME}' not found.")
            except Exception as e:
                db.session.rollback()
                print(f"An unexpected error occurred during database setup (outer exception): {e}")


# --- Utility Function: Search and Fetch from DB ---
def get_assets(query=None):
    """Fetches assets from the database based on a search query."""
    if query:
        search = f"%{query.lower()}%"

        return db.session.query(Asset).filter(
            db.or_(
                Asset.asset_type.ilike(search),
                Asset.product.ilike(search),
                Asset.name.ilike(search),
                Asset.serial_number.ilike(search),
                Asset.used_by_name.ilike(search),
                Asset.used_by_email.ilike(search),
                Asset.department.ilike(search),
                Asset.location.ilike(search)
            )
        ).all()

    return Asset.query.all()


# --- Flask Routes (UPDATED) ---
@app.route('/', methods=['GET'])
def index():
    search_query = request.args.get('query', '').strip()

    total_assets_count = Asset.query.count()

    assets = get_assets(search_query)
    result_count = len(assets)

    if assets:
        data_dicts = []
        for asset in assets:

            # Use .title() for display to capitalize words
            asset_type_display = asset.asset_type.title() if asset.asset_type else ''
            department_display = asset.department.title() if asset.department else ''
            location_display = asset.location.title() if asset.location else ''

            # Display synthetic serials clearly
            serial_display = asset.serial_number.upper() if asset.serial_number else ''
            if serial_display.startswith("SYNTHETIC_"):
                serial_display = f"{serial_display} (Non-serialized)"

            # --- NEW: Create the Edit Link ---
            edit_link = f'<a href="{url_for("edit_asset", asset_id=asset.id)}" class="btn btn-sm btn-primary">Edit</a>'

            data_dicts.append({
                'ID': asset.id,  # Include ID for debugging/reference
                'Asset Type': asset_type_display,
                'Product': asset.product or '',
                'Name': asset.name or '',
                'Serial Number': serial_display,
                'Used by (Name)': asset.used_by_name or '',
                'Used by (Email)': asset.used_by_email or '',
                'Department': department_display,
                'Location': location_display,
                'Actions': edit_link  # NEW COLUMN
            })

        results_df = pd.DataFrame(data_dicts).fillna('')

        results_html = results_df.head(100).to_html(
            classes='table table-striped table-hover table-bordered',
            index=False,
            escape=False, # <--- THIS IS THE CRITICAL FIX: Stops Pandas from escaping the HTML button code
            border=0,
        )
    else:
        results_html = f"<p class='alert alert-info'>No results found for **{search_query}**.</p>" if search_query else "<p class='alert alert-secondary'>Database is ready. Please use the search bar above.</p>"

    message = request.args.get('message')

    return render_template(
        'index.html',
        results_html=results_html,
        query=search_query,
        result_count=total_assets_count,
        message=message
    )


# --- Utility Functions for Dropdowns ---
def get_unique_asset_types():
    """Fetches a sorted list of unique asset types from the database."""
    unique_types = db.session.query(Asset.asset_type).distinct().all()

    clean_types = sorted([
        t[0].title()
        for t in unique_types
        if t[0] is not None and t[0].strip() != ''
    ])

    return clean_types


def get_unique_departments():
    """Fetches a sorted list of unique departments from the database."""
    unique_depts = db.session.query(Asset.department).distinct().all()

    clean_depts = sorted([
        d[0].title()
        for d in unique_depts
        if d[0] is not None and d[0].strip() != ''
    ])

    return clean_depts


def get_unique_locations():
    """Fetches a sorted list of unique locations from the database."""
    unique_locs = db.session.query(Asset.location).distinct().all()

    clean_locs = sorted([
        l[0].title()
        for l in unique_locs
        if l[0] is not None and l[0].strip() != ''
    ])

    return clean_locs


# --- Flask Routes (UPDATED ADD ASSET ROUTE with Dropdowns and Location Fix) ---
@app.route('/add', methods=['GET', 'POST'])
def add_asset():
    # Fetch unique lists every time the page is accessed
    unique_asset_types = get_unique_asset_types()
    unique_departments = get_unique_departments()
    unique_locations = get_unique_locations()

    # Bundle the lists into a dictionary for clean passing to the template
    template_data = {
        'unique_asset_types': unique_asset_types,
        'unique_departments': unique_departments,
        'unique_locations': unique_locations
    }

    if request.method == 'POST':
        required_fields = ['asset_type']
        for field in required_fields:
            if not request.form.get(field):
                return render_template('add_asset.html',
                                       error=f"The field '{field.replace('_', ' ').title()}' is required.",
                                       asset=request.form,
                                       **template_data)

        try:
            # --- Location Logic FIX ---
            selected_location = request.form.get('location_select')
            if selected_location == 'other':
                # Use value from the new text input if 'Other' was selected
                location = request.form.get('new_location_input', '').lower().strip()
            elif selected_location:
                # Use value from the dropdown if a valid option was selected
                location = selected_location.lower().strip()
            else:
                location = None
            # --- End Location Logic FIX ---

            # Normalize other fields
            input_serial = request.form.get('serial_number', '').lower().strip()
            serial_number = input_serial if input_serial else None

            asset_type = request.form['asset_type'].lower().strip()

            department = request.form.get('department', '').lower().strip() if request.form.get('department') else None
            used_by_email = request.form.get('used_by_email', '').lower().strip() if request.form.get(
                'used_by_email') else None

            # Check for existing serial only if one was provided
            if serial_number and Asset.query.filter_by(serial_number=serial_number).first():
                raise ValueError(
                    f"Asset with Serial Number '{request.form['serial_number'].upper()}' already exists. Please check inventory.")

            new_asset = Asset(
                asset_type=asset_type,
                product=request.form.get('product'),
                name=request.form.get('name'),
                serial_number=serial_number,
                used_by_name=request.form.get('used_by_name'),
                used_by_email=used_by_email,
                department=department,
                location=location  # Use the resolved 'location' variable
            )

            db.session.add(new_asset)
            db.session.commit()
            return redirect(url_for('index', message=f'Asset "{request.form.get("name")}" added successfully!'))
        except Exception as e:
            db.session.rollback()
            error_message = str(e).splitlines()[0]
            if 'already exists' in error_message:
                error_message = str(e)
            if 'IntegrityError' in error_message:
                error_message = "Database error: Serial Number already exists."

            return render_template('add_asset.html', error=error_message, asset=request.form, **template_data)

    # GET request: Render the form
    return render_template('add_asset.html', **template_data)


# --- Flask Routes (NEW EDIT ASSET ROUTE) ---
@app.route('/edit/<int:asset_id>', methods=['GET', 'POST'])
def edit_asset(asset_id):
    asset = Asset.query.get_or_404(asset_id)

    # Fetch unique lists for dropdowns
    unique_asset_types = get_unique_asset_types()
    unique_departments = get_unique_departments()
    unique_locations = get_unique_locations()

    template_data = {
        'unique_asset_types': unique_asset_types,
        'unique_departments': unique_departments,
        'unique_locations': unique_locations,
        'asset': asset
    }

    if request.method == 'POST':
        try:
            # --- 1. Location Logic Check ---
            selected_location = request.form.get('location_select')
            if selected_location == 'other':
                location_value = request.form.get('new_location_input', '').lower().strip()
                if not location_value:
                    raise ValueError("New location name is required if 'Other' is selected.")
            elif selected_location:
                location_value = selected_location.lower().strip()
            else:
                location_value = None

            # --- 2. Serial Number Check (Only if not synthetic) ---
            new_serial = request.form.get('serial_number', '').lower().strip() or None
            original_serial = request.form.get('original_serial_number', '').lower().strip() or None

            # If the user changed the serial number (and it's not synthetic/disabled)
            if new_serial and new_serial != original_serial:
                # Check for uniqueness against all other assets
                if Asset.query.filter(Asset.serial_number == new_serial, Asset.id != asset_id).first():
                    raise ValueError(
                        f"Serial Number '{new_serial.upper()}' already belongs to another asset. Please check inventory.")

            # --- 3. Update Asset Fields ---
            asset.asset_type = request.form['asset_type'].lower().strip()
            asset.product = request.form.get('product')
            asset.name = request.form.get('name')

            # Only update serial if it's not synthetic
            if not asset.serial_number.startswith("SYNTHETIC_"):
                asset.serial_number = new_serial

            asset.used_by_name = request.form.get('used_by_name')
            asset.used_by_email = request.form.get('used_by_email', '').lower().strip() or None
            asset.department = request.form.get('department', '').lower().strip() or None
            asset.location = location_value  # Use the resolved value

            db.session.commit()
            return redirect(url_for('index', message=f'Asset "{asset.name}" (ID: {asset_id}) updated successfully!'))

        except Exception as e:
            db.session.rollback()
            error_message = str(e).splitlines()[0]
            if 'IntegrityError' in error_message:
                error_message = "Database error: The Serial Number you entered is already in use."

            # Re-render the form with error message and form data
            asset_data = request.form.to_dict()
            asset_data['asset_type'] = asset_data.get('asset_type',
                                                      asset.asset_type)  # Preserve original values where possible

            # Since request.form returns strings, we must pass it back as a dictionary for the template to access values correctly
            # We pass the original asset object to re-populate fields on error, which contains all necessary data.
            return render_template('edit_asset.html',
                                   error=error_message,
                                   asset=asset,  # Pass original asset for context
                                   **template_data)

            # GET request: Render the form with current data
    return render_template('edit_asset.html', **template_data)

if __name__ == '__main__':
    setup_database_from_excel()
    app.run(debug=True)
