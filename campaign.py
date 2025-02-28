import pandas as pd

def process_campaign_data(file_path):
    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    
    df.columns = df.iloc[2] 
    df = df[3:].reset_index(drop=True)
    
    column_mapping = {
        "Campaign Name": "Campaign Name",
        "Send Date": "Send Date",
        "Channel": "Channel",
        "Communication Sent": "Communication Sent",
        "Delivered": "Delivered",
        "Failed DLR": "Failed DLR",
        "Open Count": "Open Count",
        "Attribution Customers": "Attribution Customers",
        "Attribution Bills": "Attribution Bills",
        "Budget": "Budget",
        "ROI": "ROI",
        "ATV": "ATV",
        "Attribution Sales": "Attribution Sales",
    }
    
    df = df[list(column_mapping.keys())].rename(columns=column_mapping)
    numeric_columns = ["Communication Sent", "Delivered", "Failed DLR", "Open Count", 
                       "Attribution Customers", "Attribution Bills", "Budget", "ROI", "ATV", "Attribution Sales"]
    df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')
    
    df.loc[df["Channel"].str.lower() == "whatsapp", "Budget"] = df["Delivered"] * 0.8
    df.loc[df["Channel"].str.lower() == "sms", "Budget"] = df["Delivered"] * 0.14
    
    whatsapp_df = df[df["Channel"].str.lower() == "whatsapp"]
    sms_df = df[df["Channel"].str.lower() == "sms"]
    
    report = {
        "Total Communication Sent (WA)": whatsapp_df["Communication Sent"].sum(),
        "Total Communication Sent (SMS)": sms_df["Communication Sent"].sum(),
        "WhatsApp % Delivery": (whatsapp_df["Delivered"].sum() / whatsapp_df["Communication Sent"].sum()) * 100 if whatsapp_df["Communication Sent"].sum() > 0 else 0,
        "SMS % Delivery": (sms_df["Delivered"].sum() / sms_df["Communication Sent"].sum()) * 100 if sms_df["Communication Sent"].sum() > 0 else 0,
        "WhatsApp % Open Count": (whatsapp_df["Open Count"].sum() / whatsapp_df["Delivered"].sum()) * 100 if whatsapp_df["Delivered"].sum() > 0 else 0,
        "WhatsApp Total Budget": whatsapp_df["Budget"].sum(),
        "SMS Total Budget": sms_df["Budget"].sum(),
        "Attribution Sales (WA)": whatsapp_df["Attribution Sales"].sum(),
        "Attribution Sales (SMS)": sms_df["Attribution Sales"].sum(),
    }
    
    output_file = "Report1.xlsx"
    df.to_excel(output_file, index=False)
    
    print("\n===== Campaign Performance Report =====")
    for key, value in report.items():
        print(f"{key}: {value:,.2f}" if isinstance(value, (int, float)) else f"{key}: {value}")
    
    return df, report, output_file

file_path = "Test1.xlsx"
df, report, output_file = process_campaign_data(file_path)


