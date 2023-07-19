# from flask import Flask, render_template, request, send_from_directory
# from werkzeug.utils import secure_filename
# from v5 import main1
# import os

# app = Flask(__name__)

# # Define the path where you want to save uploaded files
# UPLOAD_FOLDER = 'C:/Users/Vimal/OneDrive - Amrita vishwa vidyapeetham/Documents/Sem6/Ramm'
# app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# # Define the allowed extensions for the uploaded files
# ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

# def allowed_file(filename):
#     return '.' in filename and \
#            filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# @app.route('/')
# def index():
#     return render_template('indexv8_7_2_1.html')

# @app.route('/submit', methods=['POST'])
# def submit():
#     form_type = request.form.get('formType')
    
#     if form_type == 'otherForm':
#         if 'excelFile' not in request.files:
#             return 'No file part'
#         file = request.files['excelFile']

#         # Check if the file is one of the allowed types/extensions
#         if file and allowed_file(file.filename):
#             # Make the filename safe, remove unsupported chars
#             filename = secure_filename(file.filename)
#             # Move the file form the temporal folder to the upload folder we setup
#             file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

#         # Processing of the Excel file can go here...

#     elif form_type == 'templateForm':
#         data = {
#             "Teacher": str(request.form.get('teacher')),
#             "Academic_year": str(request.form.get('academicYearStart')) + "-" + str(request.form.get('academicYearEnd')),
#             "Semester": str(request.form.get('semester')),
#             "Branch": str(request.form.get('branch')),
#             "Batch": int(request.form.get('batch')),
#             "Section": str(request.form.get('section')),
#             "Subject_Code": str(request.form.get('subjectCode')),
#             "Subject_Name": str(request.form.get('subjectName')),
#             "Number_of_Students": int(request.form.get('numberOfStudents')),
#             "Number_of_COs": int(request.form.get('numberOfCOs')),
#             "Internal": float(request.form.get('internal')),
#             "External": float(request.form.get('external')),
#             "Direct": float(request.form.get('direct')),
#             "Indirect": float(request.form.get('indirect')),
#             "Default threshold %": float(request.form.get('defaultThreshold')),
#             "target": float(request.form.get('target'))
#         }
        
#         num_components = int(request.form.get('numberOfComponents'))
#         Component_Details = {}
#         for i in range(1, num_components+1):
#             Component_Details[request.form.get('componentName'+str(i))] = {"Number_of_questions": int(request.form.get('componentValue'+str(i)))}

#         filename = main1(data, Component_Details)  # Assume main1 returns the name of created file

#         # Let's assume the file is created in the same directory as your server script.
#         # Modify the directory path as necessary.
#         return send_from_directory(os.getcwd(), f"{data['Batch']}_{data['Subject_Code']}_{data['Subject_Name']}.xlsx", as_attachment=True)

#     else:
#         return 'Invalid form type'

# if __name__ == '__main__':
#     app.run(debug=True)


# pt2

# from flask import Flask, render_template, request
# from v5 import main1
# app = Flask(__name__)
# @app.route('/')
# def index():
#     return render_template('indexv8_7_2_1.html')
# @app.route('/submit', methods=['POST'])
# def submit():
#     data = {
#         # similarly for teacher]
#         "Teacher": str(request.form.get('teacher')),
#         "Academic_year": str(request.form.get('academicYearStart')) + "-" + str(request.form.get('academicYearEnd')),
#         "Semester": str(request.form.get('semester')),
#         "Branch": str(request.form.get('branch')),
#         "Batch": int(request.form.get('batch')),
#         "Section": str(request.form.get('section')),
#         "Subject_Code": str(request.form.get('subjectCode')),
#         "Subject_Name": str(request.form.get('subjectName')),
#         "Number_of_Students": int(request.form.get('numberOfStudents')),
#         "Number_of_COs": int(request.form.get('numberOfCOs')),
#         "Internal": float(request.form.get('internal')),
#         "External": float(request.form.get('external')),
#         "Direct": float(request.form.get('direct')),
#         "Indirect": float(request.form.get('indirect')),
#         "Default threshold %": float(request.form.get('defaultThreshold')),
#         "target": float(request.form.get('target'))
#         # Continue this for all your fields...
#     }
#     num_components = int(request.form.get('numberOfComponents'))
#     Component_Details = {}
#     for i in range(1, num_components+1):
#         Component_Details[request.form.get('componentName'+str(i))] = {"Number_of_questions": int(request.form.get('componentValue'+str(i)))}
#     # print("Data=", data)
#     # print("ComponentDetails=", Component_Details)
#     main1(data, Component_Details)
#     return "Data received"
# if __name__ == '__main__':

#     app.run(debug=True)


# # pt 3
# from flask import Flask, render_template, request
# from v5 import main1
# import os
# from werkzeug.utils import secure_filename

# from v3 import driver_part2


# UPLOAD_FOLDER = 'C:/Users/Vimal/OneDrive - Amrita vishwa vidyapeetham/Documents/Sem6/Ramm/Uploads'  # specify your upload folder

# app = Flask(__name__)
# app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# @app.route('/')
# def index():
#     return render_template('indexv8_7_2_1.html')

# @app.route('/submit', methods=['POST'])
# def submit():
#     # if 'file' not in request.files:
#     #     # The code for handling form data (no file)
#     #     # ...

#     #     file = request.files['file']  # Get the file from the request

#     #     if file.filename == '':
#     #         return 'No selected file.', 400

#     #     filename = secure_filename(file.filename)
#     #     file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
#     #     file.save(file_path)  # Save the file to your UPLOAD_FOLDER

#     #     # Pass the saved file's path to your driver_part2 function
#     #     return "File uploaded and passed to driver_part2 function."
#     # else:
#         # This is a submit from "Template Generation"
#         data = {
#             "Teacher": str(request.form.get('teacher')),
#             "Academic_year": str(request.form.get('academicYearStart')) + "-" + str(request.form.get('academicYearEnd')),
#             "Semester": str(request.form.get('semester')),
#             "Branch": str(request.form.get('branch')),
#             "Batch": int(request.form.get('batch')),
#             "Section": str(request.form.get('section')),
#             "Subject_Code": str(request.form.get('subjectCode')),
#             "Subject_Name": str(request.form.get('subjectName')),
#             "Number_of_Students": int(request.form.get('numberOfStudents')),
#             "Number_of_COs": int(request.form.get('numberOfCOs')),
#             "Internal": float(request.form.get('internal')),
#             "External": float(request.form.get('external')),
#             "Direct": float(request.form.get('direct')),
#             "Indirect": float(request.form.get('indirect')),
#             "Default threshold %": float(request.form.get('defaultThreshold')),
#             "target": float(request.form.get('target'))
            
#             # Continue this for all your fields...
#         }
        
#         num_components = int(request.form.get('numberOfComponents'))
#         Component_Details = {}
#         for i in range(1, num_components+1):
#             Component_Details[request.form.get('componentName'+str(i))] = {"Number_of_questions": int(request.form.get('componentValue'+str(i)))}
        
#         main1(data, Component_Details)
#         return "Data received"


# @app.route('/upload', methods=['POST'])
# def upload_file():
#     # check if the post request has the file part
#     if 'file' not in request.files:
#         return 'No file part in the request.', 400
#     file = request.files['file']
#     # if user does not select file, browser submits an empty part without filename
#     if file.filename == '':
#         return 'No selected file.', 400
#     if file:
#         filename = secure_filename(file.filename)
#         file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        
#         driver_part2("C:\\Users\\Vimal\\OneDrive - Amrita vishwa vidyapeetham\\Documents\\Sem6\\Ramm\\Uploads\\v16.xlsx")

#         return 'File successfully uploaded.', 200

# if __name__ == '__main__':
#     app.run(debug=True)
# from flask import Flask, render_template, request
# from v5 import main1
# import os
# from werkzeug.utils import secure_filename
# from v3 import driver_part2

# UPLOAD_FOLDER = 'C:/Users/Vimal/OneDrive - Amrita vishwa vidyapeetham/Documents/Sem6/Ramm/Uploads'  # specify your upload folder

# app = Flask(__name__)
# app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# @app.route('/')
# def index():
#     return render_template('index.html')

# @app.route('/details')
# def details():
#     return render_template('indexv8_7_2_1.html')

# @app.route('/submit', methods=['POST'])
# def submit():
#     data = {
#         "Teacher": str(request.form.get('teacher')),
#         "Academic_year": str(request.form.get('academicYearStart')) + "-" + str(request.form.get('academicYearEnd')),
#         "Semester": str(request.form.get('semester')),
#         "Branch": str(request.form.get('branch')),
#         "Batch": int(request.form.get('batch')),
#         "Section": str(request.form.get('section')),
#         "Subject_Code": str(request.form.get('subjectCode')),
#         "Subject_Name": str(request.form.get('subjectName')),
#         "Number_of_Students": int(request.form.get('numberOfStudents')),
#         "Number_of_COs": int(request.form.get('numberOfCOs')),
#         "Default threshold %": float(request.form.get('defaultThreshold')),
#         "target": float(request.form.get('target'))
#     }

#     # If 'Direct' is provided, calculate 'Indirect'. Similarly, if 'External' is provided, calculate 'Internal'.
#     if request.form.get('direct'):
#         data["Direct"] = float(request.form.get('direct'))
#         data["Indirect"] = 100 - data["Direct"]
#     elif request.form.get('indirect'):
#         data["Indirect"] = float(request.form.get('indirect'))
#         data["Direct"] = 100 - data["Indirect"]

#     if request.form.get('external'):
#         data["External"] = float(request.form.get('external'))
#         data["Internal"] = 100 - data["External"]
#     elif request.form.get('internal'):
#         data["Internal"] = float(request.form.get('internal'))
#         data["External"] = 100 - data["Internal"]

#     num_components = int(request.form.get('numberOfComponents'))
#     Component_Details = {}
#     for i in range(1, num_components+1):
#         Component_Details[request.form.get('componentName'+str(i))] = {"Number_of_questions": int(request.form.get('componentValue'+str(i)))}

#     main1(data, Component_Details)
#     return "Data received"

# # @app.route('/upload', methods=['POST'])
# # def upload_file():
# #     if 'file' not in request.files:
# #         return 'No file part in the request.', 400
# #     file = request.files['file']
# #     if file.filename == '':
# #         return 'No selected file.', 400
# #     if file:
# #         filename = secure_filename(file.filename)
# #         file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

# #         driver_part2("C:\\Users\\Vimal\\OneDrive - Amrita vishwa vidyapeetham\\Documents\\Sem6\\Ramm\\Uploads\\v19_1.xlsx")

# #         return 'File successfully uploaded.',
# @app.route('/upload', methods=['POST'])
# def upload_file():
#     if 'file' not in request.files:
#         return 'No file part in the request.', 400
#     file = request.files['file']
#     if file.filename == '':
#         return 'No selected file.', 400
#     if file:
#         filename = secure_filename(file.filename)
#         file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)  # full path of the file
#         file.save(file_path)  # save the file

#         driver_part2(file_path)  # pass the full path of the file to the function

#         return 'File successfully uploaded.', 200


# if __name__ == '__main__':
#     app.run(debug=True)
from flask import Flask, render_template, request, jsonify
from v5 import main1
import os
from werkzeug.utils import secure_filename
from v3 import driver_part2

UPLOAD_FOLDER = 'C:/Users/Vimal/OneDrive - Amrita vishwa vidyapeetham/Documents/Sem6/Ramm/Uploads'  # specify your upload folder

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/details')
def details():
    return render_template('indexv8_7_2_1.html')

@app.route('/submit', methods=['POST'])
def submit():
    data = {
        "Teacher": str(request.form.get('teacher')),
        "Academic_year": str(request.form.get('academicYearStart')) + "-" + str(request.form.get('academicYearEnd')),
        "Semester": str(request.form.get('semester')),
        "Branch": str(request.form.get('branch')),
        "Batch": int(request.form.get('batch')),
        "Section": str(request.form.get('section')),
        "Subject_Code": str(request.form.get('subjectCode')),
        "Subject_Name": str(request.form.get('subjectName')),
        "Number_of_Students": int(request.form.get('numberOfStudents')),
        "Number_of_COs": int(request.form.get('numberOfCOs')),
        "Default threshold %": float(request.form.get('defaultThreshold')),
        "target": float(request.form.get('target'))
    }

    direct_val = request.form.get('direct')
    external_val = request.form.get('external')

    data["Direct"] = float(direct_val) if direct_val else (100.0 - float(request.form.get('indirect')))
    data["Indirect"] = 100.0 - data["Direct"]

    data["External"] = float(external_val) if external_val else (100.0 - float(request.form.get('internal')))
    data["Internal"] = 100.0 - data["External"]

    num_components = int(request.form.get('numberOfComponents'))
    Component_Details = {}
    for i in range(1, num_components+1):
        Component_Details[request.form.get('componentName'+str(i))] = {"Number_of_questions": int(request.form.get('componentValue'+str(i)))}

    main1(data, Component_Details)

    return "Data received"
    
    # response = {
    #     "status": "Data received",
    #     "data": data,
    #     "component_details": Component_Details
    # }
    # return jsonify(response)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part in the request.', 400
    file = request.files['file']
    if file.filename == '':
        return 'No selected file.', 400
    if file:
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)  # full path of the file
        file.save(file_path)  # save the file

        driver_part2(file_path)  # pass the full path of the file to the function

        return 'File successfully uploaded.', 200


if __name__ == '__main__':
    app.run(debug=True)
