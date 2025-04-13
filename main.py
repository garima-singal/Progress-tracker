import pandas as pd
import datetime
import smtplib
from email.message import EmailMessage
import os

# ------------ CONFIGURATION ------------ #
EXCEL_FILE = 'progress_data.xlsx'
EMAIL_THRESHOLD = 70  # Percentage
SENDER_EMAIL = 'mymail@domain.com'
SENDER_PASSWORD = 'abcd1234'  # Use App password from Gmail
RECEIVER_EMAIL = 'yourmain@domain.com'
# --------------------------------------- #

# 1. Create Excel Template (if not exists)
def create_excel_template(file_path):
    if not os.path.exists(file_path):
        data = {
            'Date': pd.date_range(start='2025-04-01', periods=5, freq='D'),
            'Task': ['Task A', 'Task B', 'Task C', 'Task D', 'Task E'],
            'Target': [10, 8, 5, 12, 7],
            'Achieved': [8, 6, 5, 10, 7],
            'Remarks': ['Good', 'Needs Improvement', 'Done', 'Almost', 'Perfect']
        }
        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)
        print(f"[INFO] Excel template created: {file_path}")
    else:
        print(f"[INFO] Excel file already exists: {file_path}")

# 2. Load Data
def load_data(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"[ERROR] Failed to read Excel file: {e}")
        return None

# 3. Calculate Summary
def calculate_summary(df):
    total = len(df)
    missed = df[df['Achieved'] < df['Target']]
    missed_count = len(missed)
    completion_rate = round((1 - missed_count / total) * 100, 2) if total > 0 else 0

    summary = {
        'Total Tasks': total,
        'Missed Goals': missed_count,
        'Completion Rate': completion_rate,
        'Date': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    return summary, missed

# 4. Email Alert
def send_email_alert(summary):
    if summary['Completion Rate'] < EMAIL_THRESHOLD:
        msg = EmailMessage()
        msg.set_content(f"""
        🚨 Performance Alert 🚨

        Date: {summary['Date']}
        Completion Rate dropped to {summary['Completion Rate']}%

        Missed Goals: {summary['Missed Goals']}
        Total Tasks: {summary['Total Tasks']}
        """)
        msg['Subject'] = f'⚠️ Performance Drop Alert - {summary["Completion Rate"]}%'
        msg['From'] = SENDER_EMAIL
        msg['To'] = RECEIVER_EMAIL

        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
                smtp.send_message(msg)
            print("[INFO] Email alert sent successfully.")
        except Exception as e:
            print(f"[ERROR] Failed to send email: {e}")
    else:
        print("[INFO] Performance is good. No email sent.")

# 5. Main Function
def main():
    create_excel_template(EXCEL_FILE)
    df = load_data(EXCEL_FILE)
    if df is not None:
        summary, missed_df = calculate_summary(df)
        print("🔍 Summary:", summary)
        send_email_alert(summary)

if __name__ == "__main__":
    main()
