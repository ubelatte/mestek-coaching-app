from google.auth.transport.requests import Request
from google.auth import exceptions

# Revoke and reset credentials
credentials = None

# Assuming your service account is being loaded here
if credentials and credentials.expired and credentials.refresh_token:
    credentials.refresh(Request())
