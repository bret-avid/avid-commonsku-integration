"""
oauth_setup.py
One-time interactive OAuth flow. Opens a browser for sign-in, writes token.json.
Run on a workstation with a browser; copy token.json to the headless server.

Usage:
    python3 oauth_setup.py
"""
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]
CREDENTIALS_FILE = next(Path(__file__).parent.glob("client_secret_*.json"), None)
TOKEN_FILE = Path(__file__).parent / "token.json"

if not CREDENTIALS_FILE:
    raise SystemExit("No client_secret_*.json in this directory.")

flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
creds = flow.run_local_server(
    port=0,
    prompt="consent",
    open_browser=False,
    authorization_prompt_message=(
        "\n\n>>> OPEN THIS URL IN THE CHROME PROFILE SIGNED IN AS "
        "bret@avidapparel.ca <<<\n\n{url}\n"
    ),
)
TOKEN_FILE.write_text(creds.to_json())
print(f"Wrote {TOKEN_FILE}")
