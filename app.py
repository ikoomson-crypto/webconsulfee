from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, make_response
import sqlite3
import os
from datetime import datetime, date
from werkzeug.utils import secure_filename
import csv
import json
from io import BytesIO, StringIO
import openpyxl
from openpyxl import Workbook
import time
import base64

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['EXPORT_FOLDER'] = 'exports'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Create necessary directories
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['EXPORT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}

# Add this configuration for logo upload
app.config['LOGO_UPLOAD_FOLDER'] = 'static/uploads/logos'
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'svg'}

# Create logo upload directory
os.makedirs(app.config['LOGO_UPLOAD_FOLDER'], exist_ok=True)


def allowed_image_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_IMAGE_EXTENSIONS


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_db():
    conn = sqlite3.connect('payment_system.db')
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    db = get_db()

    # Create suppliers table
    db.execute('''
        CREATE TABLE IF NOT EXISTS suppliers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            role TEXT,
            address TEXT,
            telephone TEXT,
            bank_name TEXT,
            account_number TEXT,
            account_name TEXT,
            swift_code TEXT,
            bank_address TEXT,
            id_type TEXT,
            id_number TEXT,
            street_address TEXT,
            created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Create services table
    db.execute('''
        CREATE TABLE IF NOT EXISTS services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_id INTEGER,
            service_description TEXT NOT NULL,
            total REAL DEFAULT 0,
            status TEXT DEFAULT 'Pending',
            created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (supplier_id) REFERENCES suppliers (id)
        )
    ''')

    # Create payments table with payment_type column
    db.execute('''
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_id INTEGER,
            service_id INTEGER,
            amount REAL,
            payment_method TEXT,
            reference_number TEXT,
            payment_date DATE,
            notes TEXT,
            payment_type TEXT DEFAULT 'Earning',
            status TEXT DEFAULT 'Completed',
            created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (supplier_id) REFERENCES suppliers (id),
            FOREIGN KEY (service_id) REFERENCES services (id)
        )
    ''')

    # Create company_settings table
    db.execute('''
        CREATE TABLE IF NOT EXISTS company_settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT NOT NULL,
            company_address TEXT,
            company_phone TEXT,
            company_email TEXT,
            company_website TEXT,
            company_registration TEXT,
            currency_symbol TEXT DEFAULT '$',
            currency_code TEXT DEFAULT 'USD',
            tax_rate REAL DEFAULT 0,
            created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Check if default settings exist
    existing = db.execute('SELECT id FROM company_settings LIMIT 1').fetchone()
    if not existing:
        db.execute('''
            INSERT INTO company_settings (
                company_name, company_address, company_phone, 
                company_email, company_website, company_registration,
                currency_symbol, currency_code, tax_rate
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            'My Company',
            '123 Business Street, City, Country',
            '+1 234 567 8900',
            'info@mycompany.com',
            'www.mycompany.com',
            'REG-2024-001',
            '$',
            'USD',
            0.0
        ))

    db.commit()
    db.close()


# Initialize database
init_db()


# Company Settings Routes
@app.route('/settings', methods=['GET', 'POST'])
def settings():
    """Company settings page"""
    db = get_db()

    if request.method == 'POST':
        try:
            # Get form data
            company_name = request.form.get('company_name', '')
            company_address = request.form.get('company_address', '')
            company_phone = request.form.get('company_phone', '')
            company_email = request.form.get('company_email', '')
            company_website = request.form.get('company_website', '')
            company_registration = request.form.get('company_registration', '')
            currency_symbol = request.form.get('currency_symbol', '$')
            currency_code = request.form.get('currency_code', 'USD')
            tax_rate = float(request.form.get('tax_rate', 0))

            # Check if settings exist
            existing = db.execute('SELECT id FROM company_settings LIMIT 1').fetchone()

            if existing:
                # Update existing settings
                db.execute('''UPDATE company_settings SET 
                            company_name = ?,
                            company_address = ?,
                            company_phone = ?,
                            company_email = ?,
                            company_website = ?,
                            company_registration = ?,
                            currency_symbol = ?,
                            currency_code = ?,
                            tax_rate = ?,
                            updated_date = CURRENT_TIMESTAMP
                            WHERE id = ?''',
                           (company_name, company_address, company_phone,
                            company_email, company_website, company_registration,
                            currency_symbol, currency_code, tax_rate, existing['id']))
            else:
                # Insert new settings
                db.execute('''INSERT INTO company_settings 
                            (company_name, company_address, company_phone, 
                             company_email, company_website, company_registration,
                             currency_symbol, currency_code, tax_rate)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                           (company_name, company_address, company_phone,
                            company_email, company_website, company_registration,
                            currency_symbol, currency_code, tax_rate))

            db.commit()
            flash('Company settings updated successfully!', 'success')

        except Exception as e:
            db.rollback()
            flash(f'Error updating settings: {str(e)}', 'error')

        finally:
            db.close()

        return redirect(url_for('settings'))

    # GET - Display settings
    settings_data = db.execute('SELECT * FROM company_settings LIMIT 1').fetchone()
    db.close()

    # Get current logo
    logo_path = os.path.join('static/uploads/logos', 'company_logo.png')
    logo_exists = os.path.exists(logo_path)

    return render_template('settings.html',
                           settings=settings_data,
                           logo_exists=logo_exists)


@app.route('/settings/logo', methods=['POST'])
def upload_logo():
    """Upload company logo"""
    if 'logo' not in request.files:
        flash('No file selected', 'error')
        return redirect(url_for('settings'))

    file = request.files['logo']
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('settings'))

    if not allowed_image_file(file.filename):
        flash('Invalid file format. Please upload PNG, JPG, JPEG, GIF, or SVG', 'error')
        return redirect(url_for('settings'))

    try:
        # Secure the filename
        filename = 'company_logo.png'
        filepath = os.path.join(app.config['LOGO_UPLOAD_FOLDER'], filename)

        # Delete old logo if exists
        if os.path.exists(filepath):
            os.remove(filepath)

        # Save new logo
        file.save(filepath)
        flash('Company logo uploaded successfully!', 'success')

    except Exception as e:
        flash(f'Error uploading logo: {str(e)}', 'error')

    return redirect(url_for('settings'))


@app.route('/settings/logo/remove', methods=['POST'])
def remove_logo():
    """Remove company logo"""
    try:
        logo_path = os.path.join(app.config['LOGO_UPLOAD_FOLDER'], 'company_logo.png')
        if os.path.exists(logo_path):
            os.remove(logo_path)
            flash('Company logo removed successfully!', 'success')
        else:
            flash('No logo found to remove', 'warning')
    except Exception as e:
        flash(f'Error removing logo: {str(e)}', 'error')

    return redirect(url_for('settings'))


def get_company_settings():
    """Helper function to get company settings"""
    db = get_db()
    settings = db.execute('SELECT * FROM company_settings LIMIT 1').fetchone()
    db.close()

    if settings:
        return dict(settings)
    else:
        return {
            'company_name': 'My Company',
            'company_address': '123 Business Street, City, Country',
            'company_phone': '+1 234 567 8900',
            'company_email': 'info@mycompany.com',
            'company_website': 'www.mycompany.com',
            'company_registration': 'REG-2024-001',
            'currency_symbol': '$',
            'currency_code': 'USD',
            'tax_rate': 0.0
        }


def get_company_logo():
    """Helper function to get company logo as base64"""
    logo_path = os.path.join('static/uploads/logos', 'company_logo.png')
    if os.path.exists(logo_path):
        with open(logo_path, 'rb') as f:
            logo_data = base64.b64encode(f.read()).decode('utf-8')
            return f"data:image/png;base64,{logo_data}"
    return None


# Add these to the context processor to make settings available in all templates
@app.context_processor
def inject_company_settings():
    """Inject company settings and logo into all templates"""
    settings = get_company_settings()
    logo = get_company_logo()

    return {
        'company_settings': settings,
        'company_logo': logo
    }


@app.route('/')
def index():
    db = get_db()
    suppliers_count = db.execute('SELECT COUNT(*) as count FROM suppliers').fetchone()['count']
    payments_count = db.execute('SELECT COUNT(*) as count FROM payments').fetchone()['count']
    total_amount = db.execute('SELECT COALESCE(SUM(amount), 0) as total FROM payments').fetchone()['total']
    pending_amount = db.execute('SELECT COALESCE(SUM(total), 0) as total FROM services WHERE status = "Pending"').fetchone()['total']
    db.close()

    return render_template('index.html',
                           suppliers_count=suppliers_count,
                           payments_count=payments_count,
                           total_amount=f"{total_amount:,.0f}",
                           pending_amount=f"{pending_amount:,.0f}",
                           now=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))


@app.route('/suppliers', methods=['GET', 'POST'])
def suppliers():
    db = get_db()

    if request.method == 'POST':
        supplier_data = {
            'name': request.form['name'],
            'role': request.form.get('role', ''),
            'address': request.form['address'],
            'telephone': request.form['telephone'],
            'bank_name': request.form['bank_name'],
            'account_number': request.form['account_number'],
            'account_name': request.form['account_name'],
            'swift_code': request.form['swift_code'],
            'bank_address': request.form['bank_address'],
            'id_type': request.form['id_type'],
            'id_number': request.form['id_number'],
            'street_address': request.form['street_address']
        }

        db.execute('''INSERT INTO suppliers 
                     (name, role, address, telephone, bank_name, account_number, 
                      account_name, swift_code, bank_address, id_type, 
                      id_number, street_address)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                   (supplier_data['name'], supplier_data['role'],
                    supplier_data['address'], supplier_data['telephone'],
                    supplier_data['bank_name'], supplier_data['account_number'],
                    supplier_data['account_name'], supplier_data['swift_code'],
                    supplier_data['bank_address'], supplier_data['id_type'],
                    supplier_data['id_number'], supplier_data['street_address']))
        db.commit()
        flash('Supplier added successfully!', 'success')
        return redirect(url_for('suppliers'))

    suppliers_list = db.execute('SELECT * FROM suppliers ORDER BY id DESC').fetchall()
    db.close()
    return render_template('suppliers.html', suppliers=suppliers_list)


@app.route('/export/suppliers/<format>')
def export_suppliers(format):
    """Export suppliers to CSV, Excel, or JSON"""
    db = get_db()
    suppliers = db.execute('SELECT * FROM suppliers').fetchall()
    db.close()

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'suppliers_{timestamp}'

    if format == 'csv':
        output = StringIO()
        csv_writer = csv.writer(output)
        csv_writer.writerow(['ID', 'Name', 'Role', 'Address', 'Telephone', 'Bank Name',
                             'Account Number', 'Account Name', 'SWIFT Code',
                             'Bank Address', 'ID Type', 'ID Number',
                             'Street Address', 'Created Date'])

        for supplier in suppliers:
            csv_writer.writerow([
                supplier['id'],
                supplier['name'],
                supplier['role'] or '',
                supplier['address'] or '',
                supplier['telephone'] or '',
                supplier['bank_name'] or '',
                supplier['account_number'] or '',
                supplier['account_name'] or '',
                supplier['swift_code'] or '',
                supplier['bank_address'] or '',
                supplier['id_type'] or '',
                supplier['id_number'] or '',
                supplier['street_address'] or '',
                supplier['created_date']
            ])

        output.seek(0)
        return send_file(
            BytesIO(output.getvalue().encode('utf-8-sig')),
            mimetype='text/csv',
            as_attachment=True,
            download_name=f'{filename}.csv'
        )

    elif format == 'excel':
        wb = Workbook()
        ws = wb.active
        ws.title = "Suppliers"

        headers = ['ID', 'Name', 'Role', 'Address', 'Telephone', 'Bank Name',
                   'Account Number', 'Account Name', 'SWIFT Code',
                   'Bank Address', 'ID Type', 'ID Number',
                   'Street Address', 'Created Date']
        ws.append(headers)

        for supplier in suppliers:
            ws.append([
                supplier['id'],
                supplier['name'],
                supplier['role'] or '',
                supplier['address'] or '',
                supplier['telephone'] or '',
                supplier['bank_name'] or '',
                supplier['account_number'] or '',
                supplier['account_name'] or '',
                supplier['swift_code'] or '',
                supplier['bank_address'] or '',
                supplier['id_type'] or '',
                supplier['id_number'] or '',
                supplier['street_address'] or '',
                supplier['created_date']
            ])

        excel_file = BytesIO()
        wb.save(excel_file)
        excel_file.seek(0)

        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'{filename}.xlsx'
        )

    elif format == 'json':
        data = [dict(supplier) for supplier in suppliers]
        return send_file(
            BytesIO(json.dumps(data, indent=2, default=str).encode()),
            mimetype='application/json',
            as_attachment=True,
            download_name=f'{filename}.json'
        )

    return jsonify({'error': 'Invalid format'}), 400


@app.route('/import/suppliers', methods=['POST'])
def import_suppliers():
    """Import suppliers from CSV/Excel"""
    if 'file' not in request.files:
        flash('No file uploaded', 'error')
        return redirect(url_for('suppliers'))

    file = request.files['file']
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('suppliers'))

    if not file or not allowed_file(file.filename):
        flash('Invalid file format. Please upload .csv, .xlsx, or .xls file', 'error')
        return redirect(url_for('suppliers'))

    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    db = get_db()
    success_count = 0
    error_count = 0

    try:
        if filename.endswith('.csv'):
            with open(filepath, 'r', encoding='utf-8-sig') as csvfile:
                csv_reader = csv.DictReader(csvfile)
                for row in csv_reader:
                    try:
                        db.execute('''INSERT INTO suppliers 
                                     (name, role, address, telephone, bank_name, account_number, 
                                      account_name, swift_code, bank_address, id_type, 
                                      id_number, street_address)
                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                   (row.get('Name', ''), row.get('Role', ''),
                                    row.get('Address', ''), row.get('Telephone', ''),
                                    row.get('Bank Name', ''), row.get('Account Number', ''),
                                    row.get('Account Name', ''), row.get('SWIFT Code', ''),
                                    row.get('Bank Address', ''), row.get('ID Type', ''),
                                    row.get('ID Number', ''), row.get('Street Address', '')))
                        success_count += 1
                    except Exception as e:
                        error_count += 1
                        continue

        elif filename.endswith(('.xlsx', '.xls')):
            wb = None
            try:
                wb = openpyxl.load_workbook(filepath)
                ws = wb.active
                headers = [cell.value for cell in ws[1]]

                for row in ws.iter_rows(min_row=2, values_only=True):
                    try:
                        row_dict = dict(zip(headers, row))
                        db.execute('''INSERT INTO suppliers 
                                     (name, role, address, telephone, bank_name, account_number, 
                                      account_name, swift_code, bank_address, id_type, 
                                      id_number, street_address)
                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                   (row_dict.get('Name', ''), row_dict.get('Role', ''),
                                    row_dict.get('Address', ''), row_dict.get('Telephone', ''),
                                    row_dict.get('Bank Name', ''), row_dict.get('Account Number', ''),
                                    row_dict.get('Account Name', ''), row_dict.get('SWIFT Code', ''),
                                    row_dict.get('Bank Address', ''), row_dict.get('ID Type', ''),
                                    row_dict.get('ID Number', ''), row_dict.get('Street Address', '')))
                        success_count += 1
                    except Exception as e:
                        error_count += 1
                        continue
            finally:
                if wb:
                    wb.close()

        db.commit()

        if success_count > 0:
            flash(f'Successfully imported {success_count} suppliers!', 'success')
        if error_count > 0:
            flash(f'{error_count} records failed to import.', 'warning')

    except Exception as e:
        db.rollback()
        flash(f'Error importing file: {str(e)}', 'error')

    finally:
        db.close()
        time.sleep(0.5)
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except:
            pass

    return redirect(url_for('suppliers'))


@app.route('/supplier/<int:id>')
def view_supplier(id):
    db = get_db()
    supplier = db.execute('SELECT * FROM suppliers WHERE id = ?', (id,)).fetchone()
    db.close()
    return jsonify(dict(supplier))


@app.route('/delete_supplier/<int:id>', methods=['POST'])
def delete_supplier(id):
    """Delete a supplier and all related records"""
    db = get_db()
    try:
        db.execute('DELETE FROM payments WHERE supplier_id = ?', (id,))
        db.execute('DELETE FROM services WHERE supplier_id = ?', (id,))
        db.execute('DELETE FROM suppliers WHERE id = ?', (id,))
        db.commit()
        return jsonify({'success': True})
    except Exception as e:
        db.rollback()
        return jsonify({'success': False, 'error': str(e)})
    finally:
        db.close()


@app.route('/update_supplier', methods=['POST'])
def update_supplier():
    """Update supplier details"""
    db = get_db()

    try:
        supplier_id = request.form['id']
        supplier_data = {
            'name': request.form['name'],
            'role': request.form.get('role', ''),
            'address': request.form['address'],
            'telephone': request.form['telephone'],
            'bank_name': request.form['bank_name'],
            'account_number': request.form['account_number'],
            'account_name': request.form['account_name'],
            'swift_code': request.form['swift_code'],
            'bank_address': request.form['bank_address'],
            'id_type': request.form['id_type'],
            'id_number': request.form['id_number'],
            'street_address': request.form['street_address']
        }

        db.execute('''UPDATE suppliers SET 
                     name = ?,
                     role = ?,
                     address = ?,
                     telephone = ?,
                     bank_name = ?,
                     account_number = ?,
                     account_name = ?,
                     swift_code = ?,
                     bank_address = ?,
                     id_type = ?,
                     id_number = ?,
                     street_address = ?
                     WHERE id = ?''',
                   (supplier_data['name'], supplier_data['role'],
                    supplier_data['address'], supplier_data['telephone'],
                    supplier_data['bank_name'], supplier_data['account_number'],
                    supplier_data['account_name'], supplier_data['swift_code'],
                    supplier_data['bank_address'], supplier_data['id_type'],
                    supplier_data['id_number'], supplier_data['street_address'],
                    supplier_id))

        db.commit()
        flash('Supplier updated successfully!', 'success')

    except Exception as e:
        db.rollback()
        flash(f'Error updating supplier: {str(e)}', 'error')
        print(f"Update error: {str(e)}")

    finally:
        db.close()

    return redirect(url_for('suppliers'))


@app.route('/payments')
def payments():
    db = get_db()

    suppliers = db.execute('SELECT * FROM suppliers ORDER BY name').fetchall()

    payments = db.execute('''
        SELECT p.*, s.name as supplier_name, sv.service_description 
        FROM payments p
        JOIN suppliers s ON p.supplier_id = s.id
        JOIN services sv ON p.service_id = sv.id
        ORDER BY p.payment_date DESC
    ''').fetchall()

    pending_services = db.execute('''
        SELECT s.*, sup.name as supplier_name 
        FROM services s
        JOIN suppliers sup ON s.supplier_id = sup.id
        WHERE s.status = 'Pending'
        ORDER BY s.created_date DESC
    ''').fetchall()

    total_paid = db.execute('SELECT COALESCE(SUM(amount), 0) as total FROM payments').fetchone()['total']
    pending_amount = db.execute('SELECT COALESCE(SUM(total), 0) as total FROM services WHERE status = "Pending"').fetchone()['total']

    db.close()

    return render_template('payments.html',
                           suppliers=suppliers,
                           payments=payments,
                           pending_services=pending_services,
                           total_paid=total_paid,
                           pending_amount=pending_amount,
                           now=date.today().isoformat())


@app.route('/process_multiple_payments', methods=['POST'])
def process_multiple_payments():
    db = get_db()
    successful = 0
    failed = 0

    try:
        supplier_ids = request.form.getlist('supplier_id[]')
        service_ids = request.form.getlist('service_id[]')
        amounts = request.form.getlist('amount[]')
        payment_methods = request.form.getlist('payment_method[]')
        reference_numbers = request.form.getlist('reference_number[]')
        payment_dates = request.form.getlist('payment_date[]')
        notes_list = request.form.getlist('notes[]')
        service_descriptions = request.form.getlist('service_description[]')
        payment_types = request.form.getlist('payment_type[]')

        for i in range(len(supplier_ids)):
            try:
                if not supplier_ids[i] or not payment_methods[i]:
                    failed += 1
                    continue

                service_id = service_ids[i] if i < len(service_ids) and service_ids[i] else None
                amount = float(amounts[i]) if i < len(amounts) else 0
                service_description = service_descriptions[i] if i < len(service_descriptions) else ''
                payment_type = payment_types[i] if i < len(payment_types) else 'Earning'

                if not service_id:
                    cursor = db.execute('''INSERT INTO services (supplier_id, service_description, total, status)
                                         VALUES (?, ?, ?, ?)''',
                                        (supplier_ids[i], service_description, amount, 'Pending'))
                    service_id = cursor.lastrowid

                payment_date = payment_dates[i] if i < len(payment_dates) and payment_dates[i] else date.today().isoformat()
                reference_number = reference_numbers[i] if i < len(reference_numbers) else ''
                notes = notes_list[i] if i < len(notes_list) else ''

                db.execute('''INSERT INTO payments 
                             (supplier_id, service_id, amount, payment_method, 
                              reference_number, payment_date, notes, payment_type)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                           (supplier_ids[i], service_id, amount, payment_methods[i],
                            reference_number, payment_date, notes, payment_type))

                db.execute('UPDATE services SET status = ? WHERE id = ?', ('Paid', service_id))
                successful += 1

            except Exception as e:
                print(f"Error processing payment {i}: {str(e)}")
                failed += 1
                continue

        db.commit()

        if successful > 0:
            flash(f'Successfully processed {successful} payment(s)!', 'success')
        if failed > 0:
            flash(f'{failed} payment(s) failed to process.', 'warning')

    except Exception as e:
        db.rollback()
        flash(f'Error processing payments: {str(e)}', 'error')

    finally:
        db.close()

    return redirect(url_for('payments'))


@app.route('/get_supplier_services/<int:supplier_id>')
def get_supplier_services(supplier_id):
    db = get_db()

    services = db.execute('''
        SELECT * FROM services 
        WHERE supplier_id = ? AND status = 'Pending'
        ORDER BY created_date DESC
    ''', (supplier_id,)).fetchall()

    db.close()

    services_list = []
    for service in services:
        services_list.append({
            'id': service['id'],
            'service_description': service['service_description'],
            'amount': service['total']
        })

    return jsonify({'services': services_list})


@app.route('/get_receipt/<int:payment_id>')
def get_receipt(payment_id):
    db = get_db()

    payment = db.execute('''
        SELECT p.*, s.name as supplier_name, s.address, s.telephone,
               s.bank_name, s.account_name, s.account_number,
               sv.service_description, sv.total
        FROM payments p
        JOIN suppliers s ON p.supplier_id = s.id
        JOIN services sv ON p.service_id = sv.id
        WHERE p.id = ?
    ''', (payment_id,)).fetchone()

    db.close()

    if payment:
        receipt_html = f'''
        <div class="receipt" style="font-family: Arial, sans-serif; padding: 20px;">
            <div style="text-align: center; border-bottom: 2px solid #333; margin-bottom: 20px;">
                <h2>PAYMENT RECEIPT</h2>
                <p>Receipt No: {payment['id']:06d}</p>
                <p>Date: {payment['payment_date']}</p>
            </div>

            <div style="margin-bottom: 20px;">
                <h4>Supplier Information</h4>
                <p><strong>Name:</strong> {payment['supplier_name']}</p>
                <p><strong>Address:</strong> {payment['address'] or 'N/A'}</p>
                <p><strong>Telephone:</strong> {payment['telephone'] or 'N/A'}</p>
            </div>

            <div style="margin-bottom: 20px;">
                <h4>Payment Details</h4>
                <p><strong>Service:</strong> {payment['service_description']}</p>
                <p><strong>Amount:</strong> ${payment['amount']:,.2f}</p>
                <p><strong>Payment Method:</strong> {payment['payment_method']}</p>
                <p><strong>Reference:</strong> {payment['reference_number'] or 'N/A'}</p>
                <p><strong>Payment Type:</strong> {payment['payment_type'] or 'Earning'}</p>
            </div>

            <div style="margin-bottom: 20px;">
                <h4>Bank Details</h4>
                <p><strong>Bank Name:</strong> {payment['bank_name'] or 'N/A'}</p>
                <p><strong>Account Name:</strong> {payment['account_name'] or 'N/A'}</p>
                <p><strong>Account Number:</strong> {payment['account_number'] or 'N/A'}</p>
            </div>

            <div style="margin-top: 30px; text-align: center;">
                <p>Thank you for your payment!</p>
            </div>
        </div>
        '''
        return jsonify({'success': True, 'html': receipt_html})

    return jsonify({'success': False, 'error': 'Receipt not found'})


@app.route('/delete_payment/<int:payment_id>', methods=['POST'])
def delete_payment(payment_id):
    db = get_db()

    try:
        payment = db.execute('SELECT service_id FROM payments WHERE id = ?', (payment_id,)).fetchone()

        if payment:
            db.execute('DELETE FROM payments WHERE id = ?', (payment_id,))
            db.execute('UPDATE services SET status = ? WHERE id = ?', ('Pending', payment['service_id']))
            db.commit()
            return jsonify({'success': True})
        else:
            return jsonify({'success': False, 'error': 'Payment not found'})

    except Exception as e:
        db.rollback()
        return jsonify({'success': False, 'error': str(e)})

    finally:
        db.close()


@app.route('/download_payment_template')
def download_payment_template():
    """Download Excel template for payment import"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Payment Template"

    headers = ['Supplier Name', 'Service Description', 'Amount', 'Payment Method',
               'Reference Number', 'Payment Date', 'Payment Type', 'Notes']
    ws.append(headers)

    example_data = [
        ['ABC Company', 'Consulting Services', '500000', 'Bank Transfer',
         'TRX001', '2024-01-15', 'Earning', 'Monthly payment'],
        ['XYZ Limited', 'Web Development', '750000', 'Cash',
         '', '2024-01-20', 'Earning', 'Initial deposit'],
        ['ABC Company', 'Tax Deduction', '100000', 'Bank Transfer',
         'TAX001', '2024-01-15', 'Deduction', ''],
    ]

    for row in example_data:
        ws.append(row)

    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True)
        cell.fill = openpyxl.styles.PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.font = openpyxl.styles.Font(color="FFFFFF", bold=True)

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

    excel_file = BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)

    return send_file(
        excel_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='payment_import_template.xlsx'
    )


@app.route('/import_payments_excel', methods=['POST'])
def import_payments_excel():
    """Import multiple payments from Excel/CSV file"""
    if 'payment_file' not in request.files:
        flash('No file uploaded', 'error')
        return redirect(url_for('payments'))

    file = request.files['payment_file']
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('payments'))

    if not allowed_file(file.filename):
        flash('Invalid file format. Please upload .xlsx, .xls, or .csv file', 'error')
        return redirect(url_for('payments'))

    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    db = get_db()
    successful = 0
    failed = 0
    errors = []

    try:
        if filename.endswith('.csv'):
            with open(filepath, 'r', encoding='utf-8-sig') as csvfile:
                reader = csv.DictReader(csvfile)
                data = list(reader)
        else:
            wb = None
            try:
                wb = openpyxl.load_workbook(filepath)
                ws = wb.active
                headers = [cell.value for cell in ws[1]]
                data = []
                for row in ws.iter_rows(min_row=2, values_only=True):
                    row_dict = dict(zip(headers, row))
                    data.append(row_dict)
            finally:
                if wb:
                    wb.close()

        for idx, row in enumerate(data, start=2):
            try:
                supplier_name = str(row.get('Supplier Name', '') or row.get('supplier name', '') or '').strip()
                service_description = str(row.get('Service Description', '') or row.get('service description', '') or '').strip()
                amount = row.get('Amount', 0) or row.get('amount', 0)
                payment_method = str(row.get('Payment Method', '') or row.get('payment method', '') or '').strip()
                reference_number = str(row.get('Reference Number', '') or row.get('reference number', '') or '').strip()
                payment_date = str(row.get('Payment Date', '') or row.get('payment date', '') or '').strip()
                payment_type = str(row.get('Payment Type', '') or row.get('payment type', '') or 'Earning').strip()
                notes = str(row.get('Notes', '') or row.get('notes', '') or '').strip()

                if not supplier_name:
                    errors.append(f"Row {idx}: Supplier Name is required")
                    failed += 1
                    continue

                if not service_description:
                    errors.append(f"Row {idx}: Service Description is required")
                    failed += 1
                    continue

                if not amount or float(amount) <= 0:
                    errors.append(f"Row {idx}: Valid Amount is required")
                    failed += 1
                    continue

                if not payment_method:
                    errors.append(f"Row {idx}: Payment Method is required")
                    failed += 1
                    continue

                supplier = db.execute('SELECT id FROM suppliers WHERE LOWER(name) LIKE LOWER(?)', (f'%{supplier_name}%',)).fetchone()
                if not supplier:
                    errors.append(f"Row {idx}: Supplier '{supplier_name}' not found")
                    failed += 1
                    continue

                cursor = db.execute('''INSERT INTO services (supplier_id, service_description, total, status)
                                     VALUES (?, ?, ?, ?)''',
                                    (supplier['id'], service_description, float(amount), 'Pending'))
                service_id = cursor.lastrowid

                if not payment_date:
                    payment_date = date.today().isoformat()

                db.execute('''INSERT INTO payments 
                             (supplier_id, service_id, amount, payment_method, 
                              reference_number, payment_date, notes, payment_type)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                           (supplier['id'], service_id, float(amount), payment_method,
                            reference_number, payment_date, notes, payment_type))

                db.execute('UPDATE services SET status = ? WHERE id = ?', ('Paid', service_id))
                successful += 1

            except Exception as e:
                errors.append(f"Row {idx}: {str(e)}")
                failed += 1
                continue

        db.commit()

        if successful > 0:
            flash(f'Successfully imported {successful} payment(s)!', 'success')
        if failed > 0:
            flash(f'{failed} payment(s) failed to import.', 'warning')
        if errors:
            error_msg = '; '.join(errors[:5])
            if len(errors) > 5:
                error_msg += f' and {len(errors) - 5} more errors'
            flash(f'Errors: {error_msg}', 'error')

    except Exception as e:
        db.rollback()
        flash(f'Error processing file: {str(e)}', 'error')

    finally:
        db.close()
        time.sleep(0.5)
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except:
            pass

    return redirect(url_for('payments'))


@app.route('/bulk_payment_status')
def bulk_payment_status():
    """Get status of bulk payment imports"""
    db = get_db()

    recent_payments = db.execute('''
        SELECT COUNT(*) as count, COALESCE(SUM(amount), 0) as total
        FROM payments 
        WHERE created_date >= datetime('now', '-1 day')
    ''').fetchone()

    db.close()

    return jsonify({
        'count': recent_payments['count'],
        'total': recent_payments['total']
    })


@app.route('/payment_history')
def payment_history():
    db = get_db()

    payments_data = db.execute('''
        SELECT p.*, s.name as supplier_name, sv.service_description, p.payment_type
        FROM payments p
        JOIN suppliers s ON p.supplier_id = s.id
        JOIN services sv ON p.service_id = sv.id
        ORDER BY p.payment_date DESC
    ''').fetchall()

    payments = []
    for payment in payments_data:
        payments.append({
            'id': payment['id'],
            'payment_date': payment['payment_date'],
            'supplier_name': payment['supplier_name'],
            'service_description': payment['service_description'],
            'amount': payment['amount'],
            'payment_method': payment['payment_method'],
            'reference_number': payment['reference_number'],
            'notes': payment['notes'],
            'payment_type': payment['payment_type'] or 'Earning',
            'supplier_id': payment['supplier_id']
        })

    suppliers = db.execute('SELECT id, name FROM suppliers ORDER BY name').fetchall()
    suppliers_list = [dict(supplier) for supplier in suppliers]

    total_paid = db.execute('SELECT COALESCE(SUM(amount), 0) as total FROM payments').fetchone()['total']
    monthly_total = db.execute('SELECT COALESCE(SUM(amount), 0) as total FROM payments WHERE strftime("%Y-%m", payment_date) = strftime("%Y-%m", "now")').fetchone()['total']
    unique_suppliers = db.execute('SELECT COUNT(DISTINCT supplier_id) as count FROM payments').fetchone()['count']

    db.close()

    return render_template('payment_history.html',
                           payments=payments,
                           suppliers=suppliers_list,
                           total_paid=total_paid,
                           monthly_total=monthly_total,
                           unique_suppliers=unique_suppliers)


@app.route('/invoice/<int:payment_id>')
def invoice(payment_id):
    """Generate printable invoice for a payment"""
    db = get_db()

    payment = db.execute('''
        SELECT p.*, s.name as supplier_name, s.address, s.telephone,
               s.bank_name, s.account_name, s.account_number, s.swift_code,
               s.bank_address, s.id_type, s.id_number, s.street_address,
               sv.service_description, sv.total
        FROM payments p
        JOIN suppliers s ON p.supplier_id = s.id
        JOIN services sv ON p.service_id = sv.id
        WHERE p.id = ?
    ''', (payment_id,)).fetchone()

    db.close()

    if not payment:
        flash('Payment not found', 'error')
        return redirect(url_for('payment_history'))

    return render_template('invoice.html', payment=payment, payment_id=payment_id)


@app.route('/export_all_payments_html')
def export_all_payments_html():
    """Export all payments as HTML report"""
    db = get_db()

    payments = db.execute('''
        SELECT p.*, s.name as supplier_name, sv.service_description, p.payment_type
        FROM payments p
        JOIN suppliers s ON p.supplier_id = s.id
        JOIN services sv ON p.service_id = sv.id
        ORDER BY p.payment_date DESC
    ''').fetchall()

    total_amount = db.execute('SELECT COALESCE(SUM(amount), 0) as total FROM payments').fetchone()['total']
    db.close()

    return render_template('all_payments_report.html',
                           payments=payments,
                           total_amount=total_amount,
                           now=datetime.now())


@app.route('/payslip/<int:payment_id>')
def payslip(payment_id):
    """Generate payslip report for a payment"""
    db = get_db()

    payment = db.execute('''
        SELECT p.*, s.name as supplier_name, s.address, s.telephone,
               s.bank_name, s.account_name, s.account_number, s.swift_code,
               s.bank_address, s.id_type, s.id_number, s.street_address,
               s.role, sv.service_description, sv.total, p.payment_type
        FROM payments p
        JOIN suppliers s ON p.supplier_id = s.id
        JOIN services sv ON p.service_id = sv.id
        WHERE p.id = ?
    ''', (payment_id,)).fetchone()

    db.close()

    if not payment:
        flash('Payment not found', 'error')
        return redirect(url_for('payment_history'))

    # Format payment period (e.g., "June 2026")
    payment_period = 'N/A'
    if payment['payment_date']:
        try:
            payment_date = datetime.strptime(payment['payment_date'], '%Y-%m-%d')
            payment_period = payment_date.strftime('%B %Y')
        except:
            try:
                parts = payment['payment_date'].split('-')
                if len(parts) >= 2:
                    month_num = int(parts[1])
                    year = parts[0]
                    months = {
                        1: 'January', 2: 'February', 3: 'March', 4: 'April',
                        5: 'May', 6: 'June', 7: 'July', 8: 'August',
                        9: 'September', 10: 'October', 11: 'November', 12: 'December'
                    }
                    payment_period = f"{months.get(month_num, '')} {year}"
            except:
                payment_period = payment['payment_date']

    return render_template('payslip_report.html',
                           payment=payment,
                           payment_id=payment_id,
                           now=datetime.now(),
                           payment_period=payment_period)


@app.route('/payslip/pdf/<int:payment_id>')
def payslip_pdf(payment_id):
    """Generate payslip as PDF"""
    db = get_db()

    payment = db.execute('''
        SELECT p.*, s.name as supplier_name, s.address, s.telephone,
               s.bank_name, s.account_name, s.account_number, s.swift_code,
               s.bank_address, s.id_type, s.id_number, s.street_address,
               s.role, sv.service_description, sv.total, p.payment_type
        FROM payments p
        JOIN suppliers s ON p.supplier_id = s.id
        JOIN services sv ON p.service_id = sv.id
        WHERE p.id = ?
    ''', (payment_id,)).fetchone()

    db.close()

    if not payment:
        flash('Payment not found', 'error')
        return redirect(url_for('payment_history'))

    # Format payment period
    payment_period = 'N/A'
    if payment['payment_date']:
        try:
            payment_date = datetime.strptime(payment['payment_date'], '%Y-%m-%d')
            payment_period = payment_date.strftime('%B %Y')
        except:
            payment_period = payment['payment_date']

    # Render the payslip HTML
    html_content = render_template('payslip_pdf.html',
                                   payment=payment,
                                   payment_id=payment_id,
                                   now=datetime.now(),
                                   payment_period=payment_period)

    # Generate PDF using pdfkit
    try:
        import pdfkit
        import tempfile

        # Configure pdfkit
        wkhtmltopdf_path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
        config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

        # Create temporary HTML file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
            f.write(html_content)
            temp_html = f.name

        # Create temporary PDF file
        pdf_path = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False).name

        # Convert HTML to PDF
        pdfkit.from_file(temp_html, pdf_path, configuration=config, options={
            'page-size': 'A4',
            'encoding': 'UTF-8',
            'margin-top': '10mm',
            'margin-right': '10mm',
            'margin-bottom': '10mm',
            'margin-left': '10mm'
        })

        # Read the PDF file
        with open(pdf_path, 'rb') as f:
            pdf_data = f.read()

        # Clean up temporary files
        os.unlink(temp_html)
        os.unlink(pdf_path)

        # Send PDF as download
        response = make_response(pdf_data)
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'attachment; filename=payslip_{payment_id}_{payment["supplier_name"]}.pdf'
        return response

    except Exception as e:
        flash(f'Error generating PDF: {str(e)}', 'error')
        return redirect(url_for('payslip', payment_id=payment_id))


@app.route('/payslip/list')
def payslip_list():
    """List all payslips"""
    db = get_db()

    payments_data = db.execute('''
        SELECT p.*, s.name as supplier_name, sv.service_description, p.payment_type
        FROM payments p
        JOIN suppliers s ON p.supplier_id = s.id
        JOIN services sv ON p.service_id = sv.id
        ORDER BY p.payment_date DESC
    ''').fetchall()

    payments = []
    pay_periods_set = set()

    for payment in payments_data:
        # Format pay period (e.g., "June 2026")
        pay_period = 'N/A'
        if payment['payment_date']:
            try:
                # Handle different date formats
                payment_date_str = payment['payment_date']
                if ' ' in payment_date_str:
                    # If it has time component, split it
                    payment_date_str = payment_date_str.split(' ')[0]

                # Try to parse the date
                payment_date = datetime.strptime(payment_date_str, '%Y-%m-%d')
                pay_period = payment_date.strftime('%B %Y')
            except:
                try:
                    # Try alternative format
                    payment_date = datetime.strptime(str(payment['payment_date']), '%Y-%m-%d %H:%M:%S')
                    pay_period = payment_date.strftime('%B %Y')
                except:
                    # If all fails, use the raw value
                    pay_period = str(payment['payment_date'])[:7] if payment['payment_date'] else 'N/A'

        pay_periods_set.add(pay_period)

        payments.append({
            'id': payment['id'],
            'payment_date': payment['payment_date'],
            'pay_period': pay_period,
            'supplier_name': payment['supplier_name'],
            'service_description': payment['service_description'],
            'amount': payment['amount'],
            'payment_method': payment['payment_method'],
            'payment_type': payment['payment_type'] or 'Earning',
            'supplier_id': payment['supplier_id']
        })

    # Get suppliers for filter
    suppliers = db.execute('SELECT id, name FROM suppliers ORDER BY name').fetchall()
    suppliers_list = [dict(supplier) for supplier in suppliers]

    total_amount = db.execute('SELECT COALESCE(SUM(amount), 0) as total FROM payments').fetchone()['total']
    monthly_total = db.execute(
        'SELECT COALESCE(SUM(amount), 0) as total FROM payments WHERE strftime("%Y-%m", payment_date) = strftime("%Y-%m", "now")').fetchone()[
        'total']
    unique_suppliers = db.execute('SELECT COUNT(DISTINCT supplier_id) as count FROM payments').fetchone()['count']

    db.close()

    # Sort pay periods chronologically
    def parse_pay_period(period):
        if period == 'N/A':
            return datetime(1900, 1, 1)
        try:
            return datetime.strptime(period, '%B %Y')
        except:
            return datetime(1900, 1, 1)

    pay_periods = sorted(list(pay_periods_set), key=parse_pay_period, reverse=True)

    return render_template('payslip_list.html',
                           payments=payments,
                           pay_periods=pay_periods,
                           suppliers=suppliers_list,
                           total_amount=total_amount,
                           monthly_total=monthly_total,
                           unique_suppliers=unique_suppliers)


@app.route('/bulk_payslip_pdf', methods=['POST'])
def bulk_payslip_pdf():
    """Generate multiple payslips as PDF (zip file)"""
    import zipfile
    import tempfile

    payment_ids = request.form.getlist('payment_ids[]')

    if not payment_ids:
        flash('No payslips selected', 'error')
        return redirect(url_for('payslip_list'))

    db = get_db()
    payments = []

    for pid in payment_ids:
        payment = db.execute('''
            SELECT p.*, s.name as supplier_name, s.address, s.telephone,
                   s.bank_name, s.account_name, s.account_number, s.swift_code,
                   s.bank_address, s.id_type, s.id_number, s.street_address,
                   s.role, sv.service_description, sv.total, p.payment_type
            FROM payments p
            JOIN suppliers s ON p.supplier_id = s.id
            JOIN services sv ON p.service_id = sv.id
            WHERE p.id = ?
        ''', (pid,)).fetchone()
        if payment:
            payments.append(payment)

    db.close()

    if not payments:
        flash('No payslips found', 'error')
        return redirect(url_for('payslip_list'))

    try:
        import pdfkit
        import tempfile

        # Configure pdfkit
        wkhtmltopdf_path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
        config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

        # Create temporary directory for PDFs
        temp_dir = tempfile.mkdtemp()
        pdf_files = []

        for payment in payments:
            # Format payment period
            payment_period = 'N/A'
            if payment['payment_date']:
                try:
                    payment_date_obj = datetime.strptime(payment['payment_date'], '%Y-%m-%d')
                    payment_period = payment_date_obj.strftime('%B %Y')
                except:
                    payment_period = payment['payment_date']

            # Render HTML for each payslip
            html_content = render_template('payslip_pdf.html',
                                           payment=payment,
                                           payment_id=payment['id'],
                                           now=datetime.now(),
                                           payment_period=payment_period)

            # Generate PDF for each payslip
            pdf_filename = f"payslip_{payment['id']}_{payment['supplier_name'].replace(' ', '_')}.pdf"
            pdf_path = os.path.join(temp_dir, pdf_filename)

            # Create temporary HTML file
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html_content)
                temp_html = f.name

            # Convert HTML to PDF
            pdfkit.from_file(temp_html, pdf_path, configuration=config, options={
                'page-size': 'A4',
                'encoding': 'UTF-8',
                'margin-top': '10mm',
                'margin-right': '10mm',
                'margin-bottom': '10mm',
                'margin-left': '10mm'
            })

            # Clean up temporary HTML
            os.unlink(temp_html)
            pdf_files.append(pdf_path)

        # Create ZIP file
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for pdf_path in pdf_files:
                zip_file.write(pdf_path, os.path.basename(pdf_path))
                os.unlink(pdf_path)

        # Clean up temporary directory
        os.rmdir(temp_dir)

        # Send ZIP file
        zip_buffer.seek(0)
        response = make_response(zip_buffer.getvalue())
        response.headers['Content-Type'] = 'application/zip'
        response.headers[
            'Content-Disposition'] = f'attachment; filename=payslips_{datetime.now().strftime("%Y%m%d")}.zip'
        return response

    except Exception as e:
        flash(f'Error generating payslips: {str(e)}', 'error')
        return redirect(url_for('payslip_list'))


@app.route('/bulk_invoice_pdf', methods=['POST'])
def bulk_invoice_pdf():
    """Generate multiple invoices as PDF (zip file)"""
    import zipfile
    import tempfile

    payment_ids = request.form.getlist('payment_ids[]')

    if not payment_ids:
        flash('No payments selected', 'error')
        return redirect(url_for('payment_history'))

    db = get_db()
    payments = []

    for pid in payment_ids:
        payment = db.execute('''
            SELECT p.*, s.name as supplier_name, s.address, s.telephone,
                   s.bank_name, s.account_name, s.account_number, s.swift_code,
                   s.bank_address, s.id_type, s.id_number, s.street_address,
                   sv.service_description, sv.total
            FROM payments p
            JOIN suppliers s ON p.supplier_id = s.id
            JOIN services sv ON p.service_id = sv.id
            WHERE p.id = ?
        ''', (pid,)).fetchone()
        if payment:
            payments.append(payment)

    db.close()

    if not payments:
        flash('No invoices found', 'error')
        return redirect(url_for('payment_history'))

    try:
        import pdfkit
        import tempfile

        # Configure pdfkit
        wkhtmltopdf_path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
        config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

        # Create temporary directory for PDFs
        temp_dir = tempfile.mkdtemp()
        pdf_files = []

        for payment in payments:
            # Render HTML for each invoice
            html_content = render_template('invoice_pdf.html',
                                           payment=payment,
                                           payment_id=payment['id'])

            # Generate PDF for each invoice
            pdf_filename = f"invoice_{payment['id']}_{payment['supplier_name'].replace(' ', '_')}.pdf"
            pdf_path = os.path.join(temp_dir, pdf_filename)

            # Create temporary HTML file
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html_content)
                temp_html = f.name

            # Convert HTML to PDF
            pdfkit.from_file(temp_html, pdf_path, configuration=config, options={
                'page-size': 'A4',
                'encoding': 'UTF-8',
                'margin-top': '10mm',
                'margin-right': '10mm',
                'margin-bottom': '10mm',
                'margin-left': '10mm'
            })

            # Clean up temporary HTML
            os.unlink(temp_html)
            pdf_files.append(pdf_path)

        # Create ZIP file
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for pdf_path in pdf_files:
                zip_file.write(pdf_path, os.path.basename(pdf_path))
                os.unlink(pdf_path)

        # Clean up temporary directory
        os.rmdir(temp_dir)

        # Send ZIP file
        zip_buffer.seek(0)
        response = make_response(zip_buffer.getvalue())
        response.headers['Content-Type'] = 'application/zip'
        response.headers[
            'Content-Disposition'] = f'attachment; filename=invoices_{datetime.now().strftime("%Y%m%d")}.zip'
        return response

    except Exception as e:
        flash(f'Error generating invoices: {str(e)}', 'error')
        return redirect(url_for('payment_history'))

@app.route('/export_invoice_pdf/<int:payment_id>')
def export_invoice_pdf(payment_id):
    """Export single invoice as PDF using browser print"""
    db = get_db()

    payment = db.execute('''
        SELECT p.*, s.name as supplier_name, s.address, s.telephone,
               s.bank_name, s.account_name, s.account_number, s.swift_code,
               s.bank_address, s.id_type, s.id_number, s.street_address,
               sv.service_description, sv.total
        FROM payments p
        JOIN suppliers s ON p.supplier_id = s.id
        JOIN services sv ON p.service_id = sv.id
        WHERE p.id = ?
    ''', (payment_id,)).fetchone()

    db.close()

    if not payment:
        flash('Payment not found', 'error')
        return redirect(url_for('payment_history'))

    return render_template('invoice_pdf.html', payment=payment, payment_id=payment_id)


def amount_in_words(amount):
    """Convert numeric amount to words"""
    def number_to_words(n):
        ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine']
        teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen']
        tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']

        if n < 10:
            return ones[int(n)]
        elif n < 20:
            return teens[int(n) - 10]
        elif n < 100:
            return tens[int(n // 10)] + (' ' + ones[int(n % 10)] if n % 10 != 0 else '')
        elif n < 1000:
            return ones[int(n // 100)] + ' Hundred' + (' ' + number_to_words(n % 100) if n % 100 != 0 else '')
        elif n < 1000000:
            return number_to_words(int(n // 1000)) + ' Thousand' + (' ' + number_to_words(n % 1000) if n % 1000 != 0 else '')
        elif n < 1000000000:
            return number_to_words(int(n // 1000000)) + ' Million' + (' ' + number_to_words(n % 1000000) if n % 1000000 != 0 else '')
        else:
            return str(n)

    dollars = int(amount)
    cents = int(round((amount - dollars) * 100))

    if dollars == 0 and cents == 0:
        return "Zero"

    words = number_to_words(dollars)
    if cents > 0:
        words += f" and {cents:02d}/100"

    return words


# Register the filters with Jinja2
app.jinja_env.globals.update(amount_in_words=amount_in_words)

if __name__ == '__main__':
    app.run(debug=True)