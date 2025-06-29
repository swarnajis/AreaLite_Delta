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
        # Always save uploaded files to fixed names
        arealite_file = request.files.get("arealite_delta")
        cli_dump_file = request.files.get("cli_dump")

        if not arealite_file or not cli_dump_file:
            return "❌ Both AREALITE_Delta.txt and CLI_DUMP.txt must be uploaded.", 400

        arealite_file.save("AREALITE_Delta.txt")
        cli_dump_file.save("CLI_DUMP.txt")
        # Ensure IntTOString_Para.xlsx exists in current directory
     print("📂 Files in directory:", os.listdir())

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

#if __name__ == "__main__":
#    app.run(host="0.0.0.0", port=40000)
