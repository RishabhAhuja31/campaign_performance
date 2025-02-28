from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def process_campaign_data(file_path):
    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    
    df.columns = df.iloc[2]
    df = df[3:].reset_index(drop=True)
    
    df.loc[df["Channel"].str.lower() == "whatsapp", "Budget"] = df["Delivered"] * 0.8
    df.loc[df["Channel"].str.lower() == "sms", "Budget"] = df["Delivered"] * 0.14
    
    whatsapp_df = df[df["Channel"].str.lower() == "whatsapp"]
    sms_df = df[df["Channel"].str.lower() == "sms"]
    
    report = {
        "WhatsApp Total Budget": whatsapp_df["Budget"].sum(),
        "SMS Total Budget": sms_df["Budget"].sum(),
        "Total Communication Sent (WA)": whatsapp_df["Communication Sent"].sum(),
        "Total Communication Sent (SMS)": sms_df["Communication Sent"].sum(),
        "Attribution Sales (WA)": whatsapp_df["Attribution Sales"].sum(),
        "Attribution Sales (SMS)": sms_df["Attribution Sales"].sum(),
        "WhatsApp % Delivery": (whatsapp_df["Delivered"].sum() / whatsapp_df["Communication Sent"].sum()) * 100 if whatsapp_df["Communication Sent"].sum() > 0 else 0,
        "SMS % Delivery": (sms_df["Delivered"].sum() / sms_df["Communication Sent"].sum()) * 100 if sms_df["Communication Sent"].sum() > 0 else 0,
        "WhatsApp % Open Count": (whatsapp_df["Open Count"].sum() / whatsapp_df["Delivered"].sum()) * 100 if whatsapp_df["Delivered"].sum() > 0 else 0,
    }
    
    output_file = os.path.join(UPLOAD_FOLDER, "processed_campaign_data.xlsx")
    df.to_excel(output_file, index=False)
    
    return report, output_file

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        file = request.files["file"]
        if file:
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
            report, output_file = process_campaign_data(file_path)
            return render_template("report.html", report=report, download_link=output_file)
    return render_template("upload.html")

@app.route("/download")
def download_file():
    return send_file(os.path.join(UPLOAD_FOLDER, "processed_campaign_data.xlsx"), as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
