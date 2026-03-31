from simple_salesforce import Salesforce
from openai import OpenAI
import requests

# --- CONFIG ---
SF_USERNAME = "your_email"
SF_PASSWORD = "your_password"
SF_TOKEN = "your_security_token"

OPENAI_API_KEY = "your_openai_key"

OUTLOOK_ACCESS_TOKEN = "your_outlook_token"

# --- CONNECT TO SALESFORCE ---
sf = Salesforce(
    username=SF_USERNAME,
    password=SF_PASSWORD,
    security_token=SF_TOKEN
)

# --- GET LATEST LEAD ---
leads = sf.query("""
SELECT Id, Name, Company, Email 
FROM Lead 
WHERE OwnerId = 'YOUR_USER_ID'
ORDER BY CreatedDate DESC LIMIT 1
""")

lead = leads['records'][0]

name = lead['Name']
company = lead['Company']
email = lead['Email']

# --- GENERATE EMAIL ---
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

print(email_body)

# --- SEND EMAIL ---
url = "https://graph.microsoft.com/v1.0/me/sendMail"

headers = {
    "Authorization": f"Bearer {OUTLOOK_ACCESS_TOKEN}",
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
            {
                "emailAddress": {
                    "address": email
                }
            }
        ]
    }
}

requests.post(url, headers=headers, json=email_data)

print("Email sent!")
