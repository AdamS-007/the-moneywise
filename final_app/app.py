import pandas as pd
import re
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, jsonify
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.exc import IntegrityError

# --- Configuration & Initialization ---
app = Flask(__name__)
XLSX_FILE_NAME = 'Fresh_18.12.2025.xlsx'

# Database Configuration (SQLite)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///inventory.db'
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

    def to_dict(self):
        """Helper to convert SQL object to Dictionary for JSON API."""
        return {
            "id": self.id,
            "asset_type": self.asset_type,
            "product": self.product,
            "name": self.name,
            "serial_number": self.serial_number,
            "used_by_name": self.used_by_name,
            "used_by_email": self.used_by_email,
            "department": self.department,
            "location": self.location
        }


# --- Utility: Helper function for cleaning column headers ---
def clean_col_name(col):
    if isinstance(col, str):
        cleaned = re.sub(r'[^\w\s]', '', col).strip().replace(' ', '_')
        return cleaned.lower()
    return str(col).lower()


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


# --- Data Loading and Database Setup ---
def setup_database_from_excel():
    with app.app_context():
        db.create_all()
        if Asset.query.count() == 0:
            print(f"Populating from '{XLSX_FILE_NAME}'...")
            try:
                all_sheets = pd.read_excel(XLSX_FILE_NAME, sheet_name=None, engine='openpyxl', dtype=str)
                existing_serials = set(
                    [s[0] for s in db.session.query(Asset.serial_number).filter(Asset.serial_number.isnot(None)).all()])

                for sheet_name, df in all_sheets.items():
                    if sheet_name.lower().startswith('unnamed:') or df.empty: continue
                    df.columns = [clean_col_name(col) for col in df.columns]
                    sheet_map = {col: model_attr for model_attr, possible_cols in REVERSE_MAPPING.items() for col in
                                 df.columns if col in possible_cols}
                    sheet_asset_type = sheet_name.lower().strip()

                    for index, row in df.iterrows():
                        try:
                            asset_data = {'asset_type': sheet_asset_type}
                            for col_name, model_attr in sheet_map.items():
                                value = str(row[col_name]).strip() if pd.notnull(row[col_name]) else None
                                if value in ('nan', '', 'none'): value = None
                                if model_attr in LOWERCASE_FIELDS and value: value = value.lower()
                                asset_data[model_attr] = value

                            serial_num = asset_data.get('serial_number')
                            if not serial_num:
                                serial_num = f"SYNTHETIC_{sheet_name.replace(' ', '_').upper()}_{index + 1}"
                                asset_data['serial_number'] = serial_num

                            if serial_num in existing_serials: continue
                            db.session.add(Asset(**asset_data))
                            existing_serials.add(serial_num)
                        except Exception as e:
                            print(f"Row error: {e}")
                    db.session.commit()
            except Exception as e:
                print(f"Setup error: {e}")


# --- Helper Functions for UI ---
def get_unique_asset_types():
    return sorted([t[0].title() for t in db.session.query(Asset.asset_type).distinct() if t[0]])


def get_unique_departments():
    return sorted([d[0].title() for d in db.session.query(Asset.department).distinct() if d[0]])


def get_unique_locations():
    return sorted([l[0].title() for l in db.session.query(Asset.location).distinct() if l[0]])


# --- Flask Routes ---

@app.route('/')
def index():

    query = request.args.get('query', '').strip().lower()
    total_count = Asset.query.count()

    if query:
        search = f"%{query}%"
        assets = Asset.query.filter(db.or_(
            Asset.asset_type.ilike(search), Asset.product.ilike(search),
            Asset.name.ilike(search), Asset.serial_number.ilike(search),
            Asset.used_by_name.ilike(search), Asset.used_by_email.ilike(search),
            Asset.department.ilike(search), Asset.location.ilike(search)
        )).all()
    else:
        assets = Asset.query.limit(100).all()

    data_dicts = []
    for asset in assets:
        edit_link = f'<a href="{url_for("edit_asset", asset_id=asset.id)}" class="btn btn-sm btn-primary">Edit</a>'
        serial_display = asset.serial_number.upper()
        if serial_display.startswith("SYNTHETIC_"):
            serial_display += " (Non-serialized)"

        data_dicts.append({
            'ID': asset.id, 'Asset Type': asset.asset_type.title(), 'Product': asset.product or '',
            'Name': asset.name or '', 'Serial Number': serial_display,
            'Used by (Email)': asset.used_by_email or '', 'Location': asset.location.title() if asset.location else '',
            'Actions': edit_link
        })

    results_html = pd.DataFrame(data_dicts).to_html(
        classes='table table-hover',  # Removed 'table-striped' to let the lines shine
        index=False,
        escape=False
    ) if data_dicts else "<p>No results.</p>"
    return render_template('index.html', results_html=results_html, result_count=total_count, query=query,
                           message=request.args.get('message'))


@app.route('/add', methods=['GET', 'POST'])
def add_asset():
    template_data = {'unique_asset_types': get_unique_asset_types(), 'unique_departments': get_unique_departments(),
                     'unique_locations': get_unique_locations()}
    if request.method == 'POST':
        try:
            # Location Logic
            selected_loc = request.form.get('location_select')
            location = request.form.get('new_location_input',
                                        '').lower().strip() if selected_loc == 'other' else selected_loc.lower().strip() if selected_loc else None

            serial = request.form.get('serial_number', '').lower().strip() or None
            if serial and Asset.query.filter_by(serial_number=serial).first():
                raise ValueError("Serial Number already exists.")

            new_asset = Asset(asset_type=request.form['asset_type'].lower(), product=request.form.get('product'),
                              name=request.form.get('name'),
                              serial_number=serial, used_by_name=request.form.get('used_by_name'),
                              used_by_email=request.form.get('used_by_email', '').lower() or None,
                              department=request.form.get('department', '').lower() or None, location=location)
            db.session.add(new_asset)
            db.session.commit()
            return redirect(url_for('index', message="Asset added successfully!"))
        except Exception as e:
            db.session.rollback()
            return render_template('add_asset.html', error=str(e), **template_data)
    return render_template('add_asset.html', **template_data)


@app.route('/edit/<int:asset_id>', methods=['GET', 'POST'])
def edit_asset(asset_id):
    asset = Asset.query.get_or_404(asset_id)
    template_data = {'unique_asset_types': get_unique_asset_types(), 'unique_departments': get_unique_departments(),
                     'unique_locations': get_unique_locations(), 'asset': asset}
    if request.method == 'POST':
        try:
            selected_loc = request.form.get('location_select')
            asset.location = request.form.get('new_location_input',
                                              '').lower().strip() if selected_loc == 'other' else selected_loc.lower().strip() if selected_loc else None

            new_serial = request.form.get('serial_number', '').lower().strip() or None
            if new_serial and new_serial != asset.serial_number:
                if Asset.query.filter(Asset.serial_number == new_serial, Asset.id != asset_id).first():
                    raise ValueError("Serial number already in use.")

            asset.asset_type = request.form['asset_type'].lower()
            asset.product = request.form.get('product', '').strip() or asset.asset_type.title()
            asset.name = request.form.get('name', '').strip()
            if new_serial:
                asset.serial_number = new_serial.lower().strip()
            asset.used_by_name, asset.used_by_email = request.form.get('used_by_name'), request.form.get(
                'used_by_email', '').lower() or None
            asset.department = request.form.get('department', '').lower() or None

            db.session.commit()
            return redirect(url_for('index', message="Asset updated successfully!"))
        except Exception as e:
            db.session.rollback()
            return render_template('edit_asset.html', error=str(e), **template_data)
    return render_template('edit_asset.html', **template_data)


# --- API ENDPOINTS (CONNECTED TO DB) ---

@app.route('/api/assets', methods=['POST'])
def api_create_asset():
    data = request.get_json()
    if not data or 'serial_number' not in data: return jsonify({'error': 'Missing serial_number'}), 400
    if Asset.query.filter_by(serial_number=data['serial_number'].lower()).first(): return jsonify(
        {'error': 'Already exists'}), 409
    try:
        new_asset = Asset(asset_type=data.get('asset_type', 'unknown').lower(), product=data.get('product'),
                          name=data.get('name'),
                          serial_number=data['serial_number'].lower(),
                          used_by_email=data.get('used_by_email', '').lower() or None)
        db.session.add(new_asset)
        db.session.commit()
        return jsonify({'message': 'Created', 'id': new_asset.id}), 201
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/assets/<string:serial_number>', methods=['PUT'])
def api_update_asset(serial_number):
    data, asset = request.get_json(), Asset.query.filter_by(serial_number=serial_number.lower()).first()
    if not asset: return jsonify({'error': 'Not found'}), 404
    try:
        for field in ['product', 'name', 'used_by_name', 'used_by_email', 'department', 'location']:
            if field in data: setattr(asset, field,
                                      data[field].lower() if data[field] and field in LOWERCASE_FIELDS else data[field])
        db.session.commit()
        return jsonify({'message': 'Updated'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/assets/<string:serial_number>', methods=['GET'])
def api_get_asset(serial_number):
    asset = Asset.query.filter_by(serial_number=serial_number.lower()).first()
    return jsonify(asset.to_dict()) if asset else (jsonify({'error': 'Not found'}), 404)

## Allowing to add bulk of devices through the website

@app.route('/add-bulk', methods=['GET', 'POST'])
def add_bulk():
    template_data = {
        'unique_asset_types': get_unique_asset_types(),
        'unique_departments': get_unique_departments(),
        'unique_locations': get_unique_locations()
    }

    if request.method == 'POST':
        try:
            # 1. Get Common Data
            asset_type = request.form['asset_type'].lower()
            product = request.form.get('product')
            dept = request.form.get('department', '').lower() or None

            selected_loc = request.form.get('location_select')
            location = request.form.get('new_location_input',
                                        '').lower().strip() if selected_loc == 'other' else selected_loc.lower().strip() if selected_loc else None

            # 2. Process Serial Numbers
            raw_serials = request.form.get('serials', '')
            # Split by new line or comma, then clean up whitespace
            serial_list = [s.strip().lower() for s in re.split(r'[\n,]', raw_serials) if s.strip()]

            added_count = 0
            skipped_serials = []

            for sn in serial_list:
                # Check if serial exists
                if Asset.query.filter_by(serial_number=sn).first():
                    skipped_serials.append(sn)
                    continue

                new_asset = Asset(
                    asset_type=asset_type,
                    product=product,
                    serial_number=sn,
                    department=dept,
                    location=location
                )
                db.session.add(new_asset)
                added_count += 1

            db.session.commit()

            msg = f"Successfully added {added_count} items."
            if skipped_serials:
                msg += f" Skipped {len(skipped_serials)} duplicates."

            return redirect(url_for('index', message=msg))

        except Exception as e:
            db.session.rollback()
            return render_template('add_bulk.html', error=str(e), **template_data)

    return render_template('add_bulk.html', **template_data)


from flask import send_file
import io


@app.route('/export')
def export_assets():
    # 1. Fetch all assets from the database
    assets = Asset.query.all()

    # 2. Convert the list of objects into a list of dictionaries
    data = [asset.to_dict() for asset in assets]

    # 3. Create a Pandas DataFrame
    df = pd.DataFrame(data)

    # 4. Clean up columns for the Excel file
    # Rename internal names to friendly names
    column_mapping = {
        'id': 'ID',
        'asset_type': 'Asset Type',
        'product': 'Product/Model',
        'name': 'Internal Name',
        'serial_number': 'Serial Number',
        'used_by_name': 'Used By',
        'used_by_email': 'Email',
        'department': 'Department',
        'location': 'Location'
    }
    df = df.rename(columns=column_mapping)

    # 5. Save to an "in-memory" file (so we don't clutter your hard drive)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Inventory_Export')

    output.seek(0)

    # 6. Send the file to the browser
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Inventory_Export.xlsx'
    )

@app.route('/logs')
def view_logs():
    # Fetch all logs, newest first
    all_logs = Log.query.order_by(Log.timestamp.desc()).all()
    return render_template('logs.html', logs=all_logs)



if __name__ == '__main__':
    setup_database_from_excel()
    app.run(host='127.0.0.1', port=5019, debug=True)