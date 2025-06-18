import pandas as pd
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()

# Load Excel
df = pd.read_excel('emails.xlsx')  # Replace with your file name
df = df.dropna(subset=['From Email', 'To Email', 'Name'])

# Track current logged-in from email
current_from_email = None
server = None



# Function to convert email to env variable key format
def email_to_env_key(email):
    return email.lower().replace('@', '_').replace('.', '_')



# Email content
subject = 'Janitorial Quote - Follow up'
body = '''Hi,

Would you be interested in getting a cleaning service quote for your premises.

If you are interested, we can set up a time for a walk-through and provide an estimate.

Thanks & Regards,
Your Team
Janitorial Services - New Jersey
'''

for index, row in df.iterrows():
    from_email = row['From Email']
    to_email = row['To Email']

    if from_email != current_from_email:
        # Close previous server if open
        if server:
            server.quit()

        # Get password from .env using transformed key
        env_key = f"EMAIL_PASSWORD_{email_to_env_key(from_email)}"
        password = os.getenv(env_key)

        if not password:
            print(f"Password not found for {from_email} in .env file!")
            continue

        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(from_email, password)
            print(f"Logged in with {from_email}")
            current_from_email = from_email
        except Exception as e:
            print(f"Failed to login with {from_email}: {e}")
            continue
        
    name = row.get('Name', 'Your Team')  # Fallback if Name is missing

    personalized_body = f'''Hi,

Would you be interested in getting a cleaning service quote for your premises.

If you are interested, we can set up a time for a walk-through and provide an estimate.

Thanks & Regards,
{name}
Janitorial Services - New Jersey
    
'''
    # Compose and send email
    msg = EmailMessage()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.set_content(personalized_body)

    try:
        server.send_message(msg)
        print(f"Email sent from {from_email} to {to_email}")
    except Exception as e:
        print(f"Failed to send email from {from_email} to {to_email}: {e}")

# Close server at end
if server:
    server.quit()
