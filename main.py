import os
from simple_salesforce import Salesforce
from openai import OpenAI
from msal import ConfidentialClientApplication
import requests

# ------------------------------
# 1️⃣ Load environment variables
# ------------------------------
SF_USERNAME = os.getenv("SF_USERNAME")
SF_PASSWORD = os.getenv("SF_PASSWORD")
SF_TOKEN = os.getenv("SF_TOKEN")
SF_USER_ID = os.getenv("SF_USER_ID")  # Your Salesforce user ID

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

CLIENT_ID = os.getenv("CLIENT_ID")       # Azure App Client ID
TENANT_ID = os.getenv("TENANT_ID")       # Azure Tenant ID
CLIENT_SECRET = os.getenv("CLIENT_SECRET")  # Azure App Client Secret

# ------------------------------
# 2️⃣ Connect to Salesforce
# ------------------------------
sf = Salesforce(
    username=SF_USERNAME,
    password=SF_PASSWORD,
    security_token=SF_TOKEN
)

leads = sf.query(f"""
SELECT Id, Name, Company, Email 
FROM Lead 
WHERE OwnerId = '{SF_USER_ID}'
ORDER BY CreatedDate DESC LIMIT 1
""")

if not leads['records']:
    print("No leads found.")
    exit()

lead = leads['records'][0]
name = lead['Name']
company = lead['Company']
email = lead['Email']

# ------------------------------
# 3️⃣ Generate email using OpenAI
# ------------------------------
client = OpenAI(api_key=OPENAI_API_KEY)

prompt = f"""
Write a short, friendly sales outreach email.

Name: {name}
Company: {company}

Keep it under 100 words.
"""

response = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[{"role": "user", "content": prompt}]
)

email_body = response.choices[0].message.content
print("Generated email:\n", email_body)

# ------------------------------
# 4️⃣ Get Outlook Access Token via MSAL
# ------------------------------
app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    client_credential=CLIENT_SECRET
)

token_response = app.acquire_token_for_client(
    scopes=["https://graph.microsoft.com/.default"]
)

if "access_token" not in token_response:
    print("Error getting access token:", token_response.get("error_description"))
    exit()

access_token = token_response["access_token"]

# ------------------------------
# 5️⃣ Send email via Microsoft Graph
# ------------------------------
url = "https://graph.microsoft.com/v1.0/me/sendMail"

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

email_data = {
    "message": {
        "subject": f"Quick intro, {name}",
        "body": {
            "contentType": "Text",
            "content": email_body
        },
        "toRecipients": [
            {"emailAddress": {"address": email}}
        ]
    }
}

response = requests.post(url, headers=headers, json=email_data)

if response.status_code == 202:
    print("Email sent successfully!")
else:
    print("Failed to send email:", response.text)
