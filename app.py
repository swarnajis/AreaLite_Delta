from flask import Flask, render_template, request, send_from_directory, redirect, url_for
import os
import shutil
import pandas as pd
from AreaLite_v1_web_compatible import main

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Save uploaded files
        for key in request.files:
            file = request.files[key]
            if file.filename:
                filepath = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(filepath)
                # Move specific files to working names
                if "AREALITE_Delta" in file.filename:
                    shutil.copy(filepath, "AREALITE_Delta.txt")
                elif "CLI_DUMP" in file.filename:
                    shutil.copy(filepath, "CLI_DUMP.txt")

        # Ensure IntTOString_Para.xlsx exists in current directory
        if not os.path.exists("IntTOString_Para.xlsx"):
            return "❌ IntTOString_Para.xlsx not found in application directory.", 400

        # Run your main script logic
        try:
            main()
        except pd.errors.EmptyDataError:
            return "❌ AREALITE_Delta.txt appears to be empty or malformed. Please upload a valid file.", 400
        except Exception as e:
            return f"❌ An unexpected error occurred: {str(e)}", 500    
        # Move results to output directory
        shutil.copy("TEMP.xlsx", os.path.join(OUTPUT_FOLDER, "TEMP.xlsx"))
        shutil.copy("AREALITE_Delta_Script.txt", os.path.join(OUTPUT_FOLDER, "AREALITE_Delta_Script.txt"))

        return redirect(url_for("results"))

    return render_template("index.html")

@app.route("/results")
def results():
    return render_template("results.html")

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
