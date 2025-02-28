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
    print("Total Communication Sent (WA):", whatsapp_df["Communication Sent"].sum(), type(whatsapp_df["Communication Sent"].sum()))
    print("Total Communication Sent (SMS):", sms_df["Communication Sent"].sum(), type(sms_df["Communication Sent"].sum()))

    report = {
        "WhatsApp Total Budget": round(whatsapp_df["Budget"].sum(), 2),
        "SMS Total Budget": round(sms_df["Budget"].sum(), 2),
        "Total Communication Sent (WA)": int(whatsapp_df["Communication Sent"].sum()),
        "Total Communication Sent (SMS)": int(sms_df["Communication Sent"].sum()),
        "Attribution Sales (WA)": round(whatsapp_df["Attribution Sales"].sum(), 2),
        "Attribution Sales (SMS)": round(sms_df["Attribution Sales"].sum(), 2),
        "WhatsApp % Delivery": "{:.2f}%".format((whatsapp_df["Delivered"].sum() / whatsapp_df["Communication Sent"].sum()) * 100) if whatsapp_df["Communication Sent"].sum() > 0 else "0.00%",
        "SMS % Delivery": "{:.2f}%".format((sms_df["Delivered"].sum() / sms_df["Communication Sent"].sum()) * 100) if sms_df["Communication Sent"].sum() > 0 else "0.00%",
        "WhatsApp % Open Count": "{:.2f}%".format((whatsapp_df["Open Count"].sum() / whatsapp_df["Delivered"].sum()) * 100) if whatsapp_df["Delivered"].sum() > 0 else "0.00%",
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
