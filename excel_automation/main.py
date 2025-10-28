import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# --- 設定 ---
data_folder = "data"
output_file = "summary.xlsx"

# メール設定
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "situkut5012@gmail.com"
sender_password = "jbsm tstc twjc oksr"
receiver_email = "skoralbkw@gmail.com"

# --- Excel集計 ---
files = [f for f in os.listdir(data_folder) if f.endswith(".xlsx")]
all_data = []

for file in files:
    df = pd.read_excel(os.path.join(data_folder, file))
    df = df[['Date', 'Product', 'Sales']]
    all_data.append(df)

merged_df = pd.concat(all_data)
summary = merged_df.groupby('Date')['Sales'].sum().reset_index()
summary.to_excel(output_file, index=False)
print(f"{output_file} を作成しました。")

# --- メール送信 ---
msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = "自動集計結果"

body = "最新の売上集計結果を送付します。"
msg.attach(MIMEText(body, 'plain'))

# 添付ファイル
with open(output_file, "rb") as f:
    part = MIMEApplication(f.read(), Name=output_file)
part['Content-Disposition'] = f'attachment; filename="{output_file}"'
msg.attach(part)

# Gmail送信
try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)
    server.send_message(msg)
    server.quit()
    print("メールを送信しました。")
except Exception as e:
    print("メール送信に失敗しました:", e)
