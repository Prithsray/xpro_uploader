import os
from flask import Flask, render_template, request, jsonify,redirect,send_file,url_for
from werkzeug.utils import secure_filename
import pandas as pd
import sqlite3
from flask_cors import CORS
import requests
from requests.auth import HTTPBasicAuth
from datetime import time
import json
import xml.etree.ElementTree as ET
import time
import datetime
import io

app = Flask(__name__, static_folder='static')
CORS(app)
app.config['SECRET_KEY'] = 'your-secret-key'

# Set up the upload folder
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'downloaded_excel')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# SQLite database initialization
DATABASE =  'uploader.db'

def drop_db():
    conn = sqlite3.connect(DATABASE)
    try:
        cursor = conn.cursor()
        cursor.execute('''DROP TABLE IF EXISTS data''')
        #cursor.execute("DROP TABLE IF EXISTS backdata")
        conn.commit()
    except Exception as e:
        print(f"Error dropping table: {str(e)}")
    finally:
        conn.close()

def create_db_data():
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS data (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date DATE DEFAULT CURRENT_TIMESTAMP,
    production_order INTEGER,
    material_code VARCHAR,
    document_date VARCHAR,
    posting_date VARCHAR,
    mvt_type INTEGER,
    doc_header_text VARCHAR,
    qty VARCHAR,
    uom VARCHAR,
    plant INTEGER,
    storage_location INTEGER,
    batch VARCHAR,
    text VARCHAR,
    mfg_date VARCHAR,
    shift VARCHAR,
    in_charge VARCHAR,
    start_hrs DATETIME,
    end_hrs DATETIME,
    core_no VARCHAR,
    non_std_weight VARCHAR,
    roll_2_sigma_percent VARCHAR,
    status VARCHAR,
    downtime DATETIME,
    rejection_type VARCHAR,
    technician VARCHAR,
    child_roll_gsm VARCHAR,
    child_roll_length VARCHAR,
    output_micron VARCHAR,
    child_roll_od VARCHAR,
    no_of_joints VARCHAR,
    cus_desp VARCHAR,
    gross_weight VARCHAR,
    material_document_year VARCHAR,
    material_document VARCHAR,
    response_text TEXT,
    response_status INTEGER,
    UNIQUE(batch, posting_date, shift)
);''')
    conn.commit()
    conn.close()

def create_db_back_data():
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS backdata (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date DATE DEFAULT CURRENT_TIMESTAMP,
    production_order INTEGER,
    material_code VARCHAR,
    document_date VARCHAR,
    posting_date VARCHAR,
    mvt_type INTEGER,
    doc_header_text VARCHAR,
    qty VARCHAR,
    uom VARCHAR,
    plant INTEGER,
    storage_location INTEGER,
    batch VARCHAR,
    text VARCHAR,
    mfg_date VARCHAR,
    shift VARCHAR,
    in_charge VARCHAR,
    start_hrs DATETIME,
    end_hrs DATETIME,
    core_no VARCHAR,
    non_std_weight VARCHAR,
    roll_2_sigma_percent VARCHAR,
    status VARCHAR,
    downtime DATETIME,
    rejection_type VARCHAR,
    technician VARCHAR,
    child_roll_gsm VARCHAR,
    child_roll_length VARCHAR,
    output_micron VARCHAR,
    child_roll_od VARCHAR,
    no_of_joints VARCHAR,
    cus_desp VARCHAR,
    gross_weight VARCHAR,
    material_document_year VARCHAR,
    material_document VARCHAR,
    response_text TEXT,
    response_status INTEGER
    
);''')
    conn.commit()
    conn.close()

def convert_time(time_obj):
    if pd.isna(time_obj):
        return None
    if isinstance(time_obj,datetime.time):
        return time_obj.strftime('%H:%M:%S')  # Convert time object to string
    return str(time_obj) 


def date_to_timestamp(date_str):
    """
    Convert a date string in 'DD.MM.YYYY' format to a timestamp in milliseconds.
    """
    from datetime import datetime, timedelta
    date_format = "%d.%m.%Y"
    # Parse the date string into a datetime object
    date_time = datetime.strptime(date_str,date_format)
    # Convert the datetime object to a timestamp in milliseconds
    timestamp = int(date_time.timestamp() * 1000)
    return timestamp

def time_to_iso_duration(time_str):
    """
    Convert a time string in 'HH:MM:SS' format to ISO 8601 duration format.
    """
    from datetime import datetime, timedelta
    time_format = "%H:%M:%S"
    dt = datetime.strptime(time_str, time_format)
    delta = timedelta(hours=dt.hour, minutes=dt.minute, seconds=dt.second)
    iso_duration = f"PT{delta.seconds // 3600}H{(delta.seconds % 3600) // 60}M{delta.seconds % 60}S"
    return iso_duration

def insert_successful_data():
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    # Insert rows into backdata table based on the conditions
    cursor.execute('''
        INSERT INTO backdata (
            date, production_order, material_code, document_date, posting_date, 
            mvt_type, doc_header_text, qty, uom, plant, storage_location, batch, text, 
            mfg_date, shift, in_charge, start_hrs, end_hrs, core_no, non_std_weight, 
            roll_2_sigma_percent, status, downtime, rejection_type, technician,child_roll_gsm,child_roll_length,output_micron,
            child_roll_od,no_of_joints,cus_desp,gross_weight,material_document_year, material_document, response_text, response_status
        )
        SELECT 
            date, production_order, material_code, document_date, posting_date, 
            mvt_type, doc_header_text, qty, uom, plant, storage_location, batch, text, 
            mfg_date, shift, in_charge, start_hrs, end_hrs, core_no, non_std_weight, 
            roll_2_sigma_percent, status, downtime, rejection_type, technician,
            child_roll_gsm,child_roll_length,output_micron,child_roll_od,no_of_joints,cus_desp,gross_weight,
            material_document_year, material_document, response_text, response_status
        FROM data
        WHERE material_document IS NOT NULL
          AND material_document_year IS NOT NULL
          AND response_text = 'Data Posted Successfully'
    ''')

    conn.commit()
    conn.close()


@app.route('/')
def index():
    execute_db_functions = request.args.get('execute', 'yes')
    if execute_db_functions == 'yes':
        create_db_data()
        drop_db()
    return render_template('index.html')

@app.route('/upload', methods=['POST', 'GET'])
def upload_file():
    create_db_data()
    create_db_back_data()
    if 'excelFile' not in request.files:
        return jsonify({'error': 'No file chosen'})

    file = request.files['excelFile']
    if file.filename == '':
        return jsonify({'error': 'No selected file'})

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # Remove rows where 'Sl. No.' column is NaN
            df = df.dropna(subset=['Sl. No.'])
            
            print("DataFrame content:")
            print(df)  # Print the first few rows of the DataFrame
            message = save_to_database(df)
            #print(message)
            return jsonify({'message': message})
        except Exception as e:
            print(e)
            return jsonify({'error': f'Error processing Excel file: {str(e)}'})

    else:
        return jsonify({'error': 'Invalid file type'})


def save_to_database(df):
    count = 0
    conn = None
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        df.columns = df.columns.str.strip()
        duplicates = []
        existing_data_backdata = []
        existing_data_data = []

        expected_columns = [
            'Production Order', 'Material Code', 'DOCUMENT DATE', 'POSTING DATE', 
            'MVT. TYPE', 'DOC HEADER TEXT', 'QTY.', 'UoM', 'Plant', 
            'Storage Location', 'Batch', 'Text', 'MFG Date', 'SHIFT', 
            'IN CHARGE', 'STATRT HRS', 'END HRS', 'CORE NO', 'NON STD WEIGHT', 
            'ROLL 2 SIGMA %', 'STATUS', 'DOWNTIME', 'REJECTION TYPE', 'TECHNICIAN','CHILD ROLL GSM',
            'CHILD ROLL LENGTH','OUTPUT MICRON','CHILD ROLL OD','NUMBER OF JOINTS','CUSTOMER DESCRIPTION','GROSS WEIGHT'
        ]
        
        missing_columns = [col for col in expected_columns if col not in df.columns]
        if missing_columns:
            return f'Missing columns in uploaded file: {", ".join(missing_columns)}'

        for index, row in df.iterrows():
            count += 1
            try:
                batch = row['Batch']
                posting_date = row['POSTING DATE']
                shift = row['SHIFT']
                material_code = row['Material Code']
                mfg_date = row['MFG Date']

                # Check for existing data in backdata table
                cursor.execute('''SELECT * FROM backdata WHERE batch = ? AND posting_date = ? AND shift = ? ''',
                               (batch, posting_date, shift))
                existing_backdata = cursor.fetchone()
                if existing_backdata:
                    existing_data_backdata.append(f"Batch: {batch} - Posting Date: {posting_date} - Shift: {shift}\n ")
                    continue

                # Check for existing data in data table
                cursor.execute('''SELECT * FROM data WHERE batch = ? AND posting_date = ? AND shift = ?''',
                               (batch, posting_date, shift))
                existing_data = cursor.fetchone()
                if existing_data:
                    existing_data_data.append(f"Batch: {batch} - Posting Date: {posting_date} - Shift: {shift}\n ")
                    continue

                # Convert datetime.time objects to strings if necessary
                strt_hrs = convert_time(row['STATRT HRS'])
                end_hrs = convert_time(row['END HRS'])
                downtime = convert_time(row['DOWNTIME'])
                
                values = (
                    row['Production Order'],
                    row['Material Code'],
                    row['DOCUMENT DATE'],
                    row['POSTING DATE'],
                    row['MVT. TYPE'],
                    row['DOC HEADER TEXT'],
                    row['QTY.'],
                    row['UoM'],
                    row['Plant'],
                    row['Storage Location'],
                    row['Batch'],
                    row['Text'],
                    row['MFG Date'],
                    row['SHIFT'],
                    row['IN CHARGE'],
                    strt_hrs,
                    end_hrs,
                    row['CORE NO'],
                    row['NON STD WEIGHT'],
                    row['ROLL 2 SIGMA %'],
                    row['STATUS'],
                    downtime,
                    row['REJECTION TYPE'],
                    row['TECHNICIAN'],
                    row['CHILD ROLL GSM'],
                    row['CHILD ROLL LENGTH'],
                    row['OUTPUT MICRON'],
                    row['CHILD ROLL OD'],
                    row['NUMBER OF JOINTS'],
                    row['CUSTOMER DESCRIPTION'],
                    row['GROSS WEIGHT'],
                    None,  # Placeholder for MaterialDocumentYear
                    None   # Placeholder for MaterialDocument
                )

                cursor.execute('''INSERT INTO data (
                                    production_order,
                                    material_code,
                                    document_date,
                                    posting_date,
                                    mvt_type,
                                    doc_header_text,
                                    qty,
                                    uom,
                                    plant,
                                    storage_location,
                                    batch,
                                    text,
                                    mfg_date,
                                    shift,
                                    in_charge,
                                    start_hrs,
                                    end_hrs,
                                    core_no,
                                    non_std_weight,
                                    roll_2_sigma_percent,
                                    status,
                                    downtime,
                                    rejection_type,
                                    technician,
                                    child_roll_gsm,
                                    child_roll_length,
                                    output_micron,
                                    child_roll_od,
                                    no_of_joints,
                                    cus_desp,
                                    gross_weight,
                                    material_document_year,
                                    material_document
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?,?,?,?,?)''', values)
            except Exception as e:
                print(f"Error inserting row {index}: {e}")

        conn.commit()
        
        message1 = ""
        message2 = ""
        if existing_data_backdata:
            message1 += f"Skipping the following data as GRN already created for the following data:\n\n{' '.join(existing_data_backdata)}\n"
        if existing_data_data:
            message2 +=  f"You have already uploaded the following data once:\n\n{''.join(existing_data_data)}\n"
        
        message3 = "\n"
        
        

        #print(f"Message content: '{message}' fsfsf")  # Debugging print statement

        if message1 or message2 :
            message = message1 + message3 + message2
            return message
        else:
            return "File uploaded and data saved successfully"

    except sqlite3.Error as e:
        return f"Database error: {e}"
    except Exception as e:
        return f"Unexpected error: {e}"
    finally:
        if conn:
            conn.close()



@app.route('/data', methods=['GET', 'POST'])
def get_data():
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM data")
    rows = cursor.fetchall()
    conn.close()
    return jsonify(rows)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}


def create_json_data(data_rows):
    # Convert dates and times to the required formats
    document_date = date_to_timestamp(data_rows[0]['document_date'])
    posting_date = date_to_timestamp(data_rows[0]['posting_date'])
    mfg_date = date_to_timestamp(data_rows[0]['date_of_manufacture'])
    
    json_data = {
        "DocumentDate": f"/Date({document_date})/",
        "PostingDate": f"/Date({posting_date})/",
        "CreatedByUser": "",
        "MaterialDocumentHeaderText": data_rows[0]['doc_header_text'],
        "GoodsMovementCode": "02",
        "to_MaterialDocumentItem": {
            "results": []
        }
    }
    count=1
    for data_row in data_rows:
        starttime = time_to_iso_duration(data_row['start_time'])
        endtime = time_to_iso_duration(data_row['end_time'])
        downtime = time_to_iso_duration(data_row['downtime'])
        
        item = {
            "MaterialDocumentItem": "",
            "Material": "",
            "Plant": f"{abs(data_row['plant'])}",
            "StorageLocation": f"{data_row['storage_location']}",
            "Batch": f"{data_row['batch']}",
            "GoodsMovementType": f"{data_row['mvt_type']}",
            "InventoryStockType": "",
            "InventoryValuationType": "",
            "InventorySpecialStockType": "",
            "Supplier": "",
            "Customer": "",
            "SalesOrder": "",
            "SalesOrderItem": "",
            "SalesOrderScheduleLine": "",
            "PurchaseOrder": "",
            "PurchaseOrderItem": "",
            "WBSElement": "",
            "ManufacturingOrder": f"{data_row['production_odr']}",
            "ManufacturingOrderItem": "",
            "GoodsMovementRefDocType": "F",
            "GoodsMovementReasonCode": "",
            "Delivery": "",
            "DeliveryItem": "",
            "AccountAssignmentCategory": "",
            "CostCenter": "",
            "ControllingArea": "",
            "CostObject": "",
            "GLAccount": "",
            "FunctionalArea": "",
            "ProfitabilitySegment": "",
            "ProfitCenter": "",
            "MasterFixedAsset": "",
            "FixedAsset": "",
            "MaterialBaseUnit": "",
            "QuantityInBaseUnit": "0",
            "EntryUnit": f"{data_row['uom']}",
            "QuantityInEntryUnit": f"{data_row['quantity']}",
            "MaterialDocumentItemText": f"{data_row['text']}",
            "ShelfLifeExpirationDate": f"/Date({mfg_date})/",
            "ManufactureDate": f"/Date({mfg_date})/",
            "SerialNumbersAreCreatedAutomly": True,
            "Reservation": "",
            "ReservationItem": "",
            "ReservationItemRecordType": "",
            "ReservationIsFinallyIssued": True,
            "SpecialStockIdfgSalesOrder": "",
            "SpecialStockIdfgSalesOrderItem": "",
            "SpecialStockIdfgWBSElement": "",
            "IsAutomaticallyCreated": "",
            "MaterialDocumentLine": f"{count}",
            "MaterialDocumentParentLine": "",
            "HierarchyNodeLevel": "",
            "GoodsMovementIsCancelled": False,
            "ReversedMaterialDocumentYear": "",
            "ReversedMaterialDocument": "",
            "ReversedMaterialDocumentItem": "",
            "ReferenceDocumentFiscalYear": "",
            "InvtryMgmtRefDocumentItem": "",
            "InvtryMgmtReferenceDocument": "",
            "MaterialDocumentPostingType": "",
            "InventoryUsabilityCode": "",
            "EWMWarehouse": "",
            "EWMStorageBin": "",
            "DebitCreditCode": "",
            "YY1_RejectionType_MMI": f"{data_row['rejection_type']}",
            "YY1_Roll2Sigma_MMI": f"{data_row['roll_2_sigma_percent']}",
            "YY1_ShiftInCharge_MMI": "prithwish",
            "YY1_Technician_MMI": f"{data_row['technician']}",
            "YY1_NSTD_MMI": f"{data_row['non_std_weight']}",
            "YY1_DownTimeMin_MMI": f"{downtime}",
            "YY1_StartHrs1_MMI": f"{starttime}",
            "YY1_FinishHrs_MMI": f"{endtime}",
            "YY1_SHIFT_MIGO_MMI": f"{data_row['shift']}",
            "YY1_CoreNo_MMI": f"{data_row['core_no']}",
            "YY1_STATUS_MIGO_MMI": f"{data_row['status']}",
            "YY1_CUSTOMER_NAME_MMI": f"{data_row['cus_desp']}",
            "YY1_NO_OF_JOINTS_MMI": f"{data_row['no_of_joints']}",
            "YY1_OUTSIDE_DIA_MMI": f"{data_row['child_roll_od']}",
            "YY1_OUTPUT_MICRON_MMI": f"{data_row['output_micron']}",
            "YY1_Roll_GSM_MMI": f"{data_row['child_roll_gsm']}",
            "YY1_ROLL_LENGTH_MMI": f"{data_row['child_roll_length']}",
            "YY1_GROSS_WT_MMI": f"{data_row['gross_weight']}",
            "to_SerialNumbers": {
                "results": [{
                    "Material": "",
                    "SerialNumber": "",
                    "MaterialDocument": "",
                    "MaterialDocumentItem": "",
                    "MaterialDocumentYear": "",
                    "ManufacturerSerialNumber": ""
                }]
            }
        }
        count=count+1
        json_data["to_MaterialDocumentItem"]["results"].append(item)

    return json_data

def post_json_to_sap(json_data):
    odata_endpoint = "https://my411220-api.s4hana.cloud.sap/sap/opu/odata/sap/API_MATERIAL_DOCUMENT_SRV/A_MaterialDocumentHeader/"

    sap_username = "API_USER"
    sap_password = "rPonb9AqaUQxtxfGNUYLRnXAmhar@AgNMfvtCwes"

    headers_token = {
        'x-CSRF-Token': 'fetch',
    }

    auth = HTTPBasicAuth(sap_username, sap_password)
    response_token = requests.get(odata_endpoint, headers=headers_token, auth=auth)

    csrf_token = response_token.headers.get('x-csrf-token')
    cookie_string = response_token.headers.get('set-cookie')

    cookie_pairs = cookie_string.split(';')
    session_id = None
    for pair in cookie_pairs:
        if "sap-XSRF_CTW_100" in pair:
            session_id = pair.split('=')[2]
            break

    if not csrf_token or not session_id:
        raise Exception("Failed to fetch CSRF token or session ID")

    cookies = {
        'sap-usercontext': 'sap-client=100',
        'sap-XSRF_CTW_100': session_id
    }

    headers_post = {
        'Content-Type': 'application/json',
        'x-csrf-Token': csrf_token,
    }

    response_post = requests.post(odata_endpoint, headers=headers_post, auth=auth, cookies=cookies, data=json.dumps(json_data))

    material_document_year = None
    material_document = None
    response_text = response_post.text

    print(response_text)
    if response_post.status_code == 201:
        try:
            # Parse XML response for success
            root = ET.fromstring(response_text)
            namespace = {'d': 'http://schemas.microsoft.com/ado/2007/08/dataservices'}
            material_document_year = root.find('.//d:MaterialDocumentYear', namespace).text
            material_document = root.find('.//d:MaterialDocument', namespace).text
            response_text = "Data Posted Successfully"
        except ET.ParseError:
            response_text = "Failed to parse XML response."
        except AttributeError:
            response_text = "Failed to extract data from XML response."
    else:
        try:
            # Parse XML response for error
            root = ET.fromstring(response_text)
            namespace = {'': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'}
            error_messages = root.findall('.//errordetail/message', namespace)
            
            # Save only the first error message
            if error_messages:
                response_text = "Error: " + error_messages[0].text
            else:
                response_text = "Unknown error occurred."
        except ET.ParseError:
            response_text = "Failed to parse error response."
        except AttributeError:
            response_text = "Failed to extract error messages from XML response."

    return response_post.status_code, response_text, material_document_year, material_document


@app.route('/process_grn', methods=['GET', 'POST'])
def create_grn():
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    # Fetch rows where material_document is NULL
    cursor.execute("SELECT * FROM data WHERE material_document IS NULL")
    rows = cursor.fetchall()
    print(rows)

    # Group rows by production_order
    production_orders = {}
    for row in rows:
        (id, date, production_order, material_code, document_date, posting_date, mvt_type, doc_header_text, qty, uom, plant, storage_location, batch, text, mfg_date, shift, in_charge, start_hrs, end_hrs, core_no, non_std_weight, roll_2_sigma_percent, status, downtime, rejection_type, technician,child_roll_gsm,child_roll_length,output_micron,child_roll_od,no_of_joints,cus_desp, material_document_year, material_document, response_text, response_status) = row
        if production_order not in production_orders:
            production_orders[production_order] = []
        production_orders[production_order].append({
            'document_date': document_date,
            'posting_date': posting_date,
            'mvt_type': mvt_type,
            'doc_header_text': doc_header_text,
            'storage_location': storage_location,
            'batch': batch,
            'quantity': qty,
            'uom': uom,
            'date_of_manufacture': mfg_date,
            'shift': shift,
            'material_code': material_code,
            'plant': plant,
            'text': text,
            'core_no': core_no,
            'non_std_weight': non_std_weight,
            'roll_2_sigma_percent': roll_2_sigma_percent,
            'status': status,
            'downtime': downtime,
            'rejection_type': rejection_type,
            'technician': technician,
            'start_time': start_hrs,
            'end_time': end_hrs,
            'production_odr': production_order,
            'child_roll_gsm':child_roll_gsm,
            'child_roll_length':child_roll_length,
            'output_micron':output_micron,
            'child_roll_od':child_roll_od,
            'no_of_joints':no_of_joints,
            'cus_desp':cus_desp,
            'gross_weight':gross_weight


        })

    for production_order, data_rows in production_orders.items():
        # Create JSON data
        json_data = create_json_data(data_rows)

        # Post data to SAP and get the response
        status_code, response_text, material_document_year, material_document = post_json_to_sap(json_data)
        print(status_code, response_text, material_document_year, material_document)

        # Update the database with the response
        for data_row in data_rows:
            cursor.execute('''UPDATE data 
                              SET response_text = CASE WHEN response_text IS NULL OR response_text = '' THEN ? ELSE response_text END,
                                  response_status = CASE WHEN response_status IS NULL OR response_status = '' THEN ? ELSE response_status END,
                                  material_document_year = CASE WHEN material_document_year IS NULL OR material_document_year = '' THEN ? ELSE material_document_year END,
                                  material_document = CASE WHEN material_document IS NULL OR material_document = '' THEN ? ELSE material_document END
                              WHERE production_order = ?''',
                           (response_text, status_code, material_document_year, material_document, production_order))

    conn.commit()
    conn.close()
    insert_successful_data()
    return redirect(url_for('index', execute='no'))


@app.route('/download_table', methods=['GET'])
def download_table_as_excel():
    # Connect to the SQLite database
    conn = sqlite3.connect(DATABASE)
    
    try:
        # Query the database to get the data from the specified table
        query = f"SELECT * FROM data"
        df = pd.read_sql_query(query, conn)
        
        # Convert the DataFrame to an Excel file in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        output.seek(0)

        # Drop the table after sending the file
        drop_db()

        # Send the Excel file as a response
        return send_file(
            output,
            download_name=f"report.xlsx",
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)})
    finally:
        conn.close()

@app.route('/deletebothdb', methods=['GET'])
def drop_both_db():
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute("DROP TABLE IF EXISTS data")
    cursor.execute("DROP TABLE IF EXISTS backdata")
    conn.commit()
    return redirect(url_for('index', execute='no'))
  

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
