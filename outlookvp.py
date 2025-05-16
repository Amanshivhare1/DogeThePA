import pandas as pd
import os
import subprocess

# --- CONFIG ---
EMAIL_SENDER = "aanssha@theloversai.co.in"
ATTACHMENT_PATH = "test.pdf"
EXCEL_PATH = "contacts.xlsx"
NAME_COLUMN = "Name"
EMAIL_COLUMN = "Email"
STATUS_COLUMN = "Status"

# --- Email Templates ---
SUBJECT_TEMPLATE = "Wedding Tech Platform Startup - LoversAI !!"
BODY_TEMPLATE = """Hi {name},

We are a wedding tech platform startup — aiming to build the biggest monopolistic company in the Indian wedding industry, a $130 billion market.

We are creating a wed-tech platform for vendors and couples to connect, with AI to visualize weddings beforehand. We’re also offering SaaS tools for vendors to manage and redesign their inventory, and providing services for rituals and astrological needs. We're building a dark store to store all this material, similar to Zepto, and an app to order all fashion jewelry — from big brands like Sabyasachi and Manish Malhotra, to local giants in Chandni Chowk and Sadar Bazaar — which are the key markets for budget wedding shopping across India. Globally, we plan to expand by introducing music festivals named “Lovers Card” and entering the matchmaking and dating industry.

I truly and genuinely require your help — more than just money — because your guidance can turn our monopolistic vision into a reality, and make it happen faster. I’m looking for the kind of support VCs offer in their earliest days.

I also got into Shark Tank India Season 4 (first round) and helped design the very famous Ambani wedding with the main wedding planner. There is still no major company dominating this massive market in India.

Short pitch deck – 1–3 year plan – short pitch deck
Long pitch deck – 7–10 year plan – long pitch deck

With fingers crossed,
Aanssha
LoversAI – The Wedding Tech Platform Startup
www.loversai.co.in
https://www.linkedin.com/in/aanssha/
+91 98216 40951
"""

# --- Load Contacts ---
try:
    df = pd.read_excel(EXCEL_PATH)
    if STATUS_COLUMN not in df.columns:
        df[STATUS_COLUMN] = ""
except FileNotFoundError:
    print(f"❌ Excel file not found: {EXCEL_PATH}")
    exit()

# --- Check Attachment ---
if not os.path.exists(ATTACHMENT_PATH):
    print(f"❌ Attachment not found: {ATTACHMENT_PATH}")
    exit()

# --- Send Emails via AppleScript + Outlook ---
for idx, row in df.iterrows():
    name = row[NAME_COLUMN]
    email = row[EMAIL_COLUMN]

    subject = SUBJECT_TEMPLATE.format(name=name).replace('"', '\\"')
    body = BODY_TEMPLATE.format(name=name).replace('"', '\\"').replace('\n', '\\n')
    attachment_path = os.path.abspath(ATTACHMENT_PATH)

    applescript = f'''
    tell application "Microsoft Outlook"
        set newMessage to make new outgoing message
        tell newMessage
            set subject to "{subject}"
            set content to "{body}"
            make new recipient at newMessage with properties {{email address:{{address:"{email}"}}}}
            make new attachment with properties {{file:POSIX file "{attachment_path}"}}
            send
        end tell
    end tell
    '''

    process = subprocess.Popen(['osascript', '-'], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate(applescript.encode('utf-8'))

    if stderr:
        print(f"❌ Failed to send email to {name} <{email}>: {stderr.decode('utf-8')}")
        df.loc[idx, STATUS_COLUMN] = "Failed"
    else:
        print(f"✅ Email sent to {name} <{email}>")
        df.loc[idx, STATUS_COLUMN] = "Sent"

# --- Save Updated Status to Excel ---
df.to_excel(EXCEL_PATH, index=False)
print("\n🎉 All emails processed!")
