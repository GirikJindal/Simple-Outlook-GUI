# Simple Outlook GUI

A tkinter-based Outlook client that connects to Microsoft 365 / Outlook accounts using the Microsoft Graph API. This application exposes full Outlook functionality through a desktop GUI — including reading, composing, drafting, sending and deleting mail; folder and mailbox management; calendar viewing and editing; viewing/downloading attachments; and other Outlook features — without embedding user credentials in the app (authentication is performed via Microsoft identity platform / OAuth2).

Table of contents
- Features
- How it works
- Requirements
- Azure App Registration (quick)
- Configuration
- Installation
- Running the app
- Security considerations
- Troubleshooting
- Contributing
- License

## Features
- Full email lifecycle: read, send, reply, forward, draft, delete
- Folder management: list mail folders, create/rename/delete folders, move messages
- Mail search and filtering
- Attachment viewing and downloading
- Calendar: list, create, update, delete events; view free/busy
- Contacts: list and basic contact management (where supported)
- Multi-account support (via separate Azure app consent or multiple sign-ins)
- Token caching and refresh using Microsoft Authentication Library (MSAL)
- Optional offline access scopes for long-lived refresh tokens (subject to tenant policy)

## How it works (high level)
1. The app authenticates users with Microsoft identity platform (Azure AD) using OAuth2.
2. After sign-in and consent, the app obtains access tokens for Microsoft Graph.
3. The app calls Graph endpoints (e.g., /me/mailFolders, /me/messages, /me/calendar/events, /me/contacts) to read and write mailbox and calendar data.
4. Tokens are cached and refreshed securely using MSAL; long-lived refresh flows depend on tenant policies.
5. The GUI exposes Graph-backed operations so users can operate their mailbox through a local Tkinter interface.

## Requirements
- Python 3.8 or later
- A registered application in Azure AD (see Azure App Registration section)
- Packages:
  - msal
  - requests
  - msal-extensions (optional, for secure token caching)
  - tkinter (standard on many Python installs; on some OSes requires separate install)
  - other dependencies listed in requirements.txt (if provided)

Example pip install:
```bash
pip install msal requests msal-extensions
```

## Azure App Registration (quick)
To allow the app to access Microsoft Graph for your tenant/account:

1. Go to the Azure Portal -> Azure Active Directory -> App registrations -> New registration.
2. Set a name (e.g., Simple-Outlook-GUI). For platform add "Mobile and desktop applications" if you use a redirect URI like `http://localhost`.
3. Note the Application (client) ID and Directory (tenant) ID.
4. Under "Authentication" add a redirect URI if you're using the interactive flow (e.g., `http://localhost`).
5. Under "API permissions" add delegated permissions for Microsoft Graph:
   - Mail.Read, Mail.ReadWrite, Mail.Send, Mail.ReadWrite.Shared
   - MailboxSettings.ReadWrite (if you need mailbox settings)
   - Calendars.ReadWrite
   - Contacts.ReadWrite
   - Files.Read (if accessing attachments stored in OneDrive/SharePoint)
   - offline_access and openid, profile, email (for sign-in and refresh)
6. Click "Grant admin consent" if your tenant requires admin consent for the requested scopes or ask users to consent individually.
7. If you plan to use client credential flow (app-only) you must create a client secret/certificate and grant application permissions instead.

Important: Choose the least-privilege set of permissions necessary for your scenario.

## Configuration
The app expects configuration for the Azure AD app and the scopes to request. Typical configuration options:
- CLIENT_ID (Application / client ID)
- TENANT_ID (Directory / tenant ID)
- CLIENT_SECRET (only for app-only/service flows; do NOT embed in distributed builds)
- REDIRECT_URI (e.g., http://localhost)
- SCOPE (e.g., ["User.Read","Mail.ReadWrite","Mail.Send","Calendars.ReadWrite","offline_access"])

You can supply configuration via:
- A config file (e.g., config.json)
- Environment variables (recommended for secrets)
- An interactive prompt (for development)

Example config.json:
```json
{
  "client_id": "YOUR_CLIENT_ID",
  "tenant_id": "YOUR_TENANT_ID",
  "redirect_uri": "http://localhost",
  "scopes": ["openid","profile","offline_access","Mail.ReadWrite","Mail.Send","Calendars.ReadWrite","Contacts.ReadWrite"]
}
```

## Installation
1. Clone the repository:
```bash
git clone https://github.com/GirikJindal/Simple-Outlook-GUI.git
cd Simple-Outlook-GUI
```
2. Create a Python virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate   # macOS/Linux
venv\Scripts\activate      # Windows
```
3. Install dependencies:
```bash
pip install -r requirements.txt
# or
pip install msal requests msal-extensions
```

## Running the app
1. Ensure your configuration (client id, tenant id, redirect_uri, scopes) is set (env vars or config file).
2. Start the GUI:
```bash
python main.py
```
3. On first run the app will open a browser for sign-in/consent (interactive OAuth). After successful authentication, the app will cache tokens and allow immediate use of mailbox and calendar features.

If the repo uses a different entrypoint, replace `main.py` with the relevant script.

## Security considerations
- Do NOT commit client secrets, certificates, or tokens to source control.
- Use environment variables or secure OS-level credential stores for secrets.
- Use the delegated user flow where possible to avoid storing long-lived credentials.
- Respect tenant admin policies: some tenants block offline_access or certain scopes.
- Only request the minimal scopes required for your users' scenarios.
- When distributing builds to other users, avoid bundling app secrets — use interactive or device code flows that do not require embedding a secret.

## Troubleshooting
- "Consent required" or "Insufficient privileges": ensure the requested Graph scopes are granted by the tenant admin.
- Authentication fails on redirect: confirm redirect URI in Azure matches the app's redirect.
- Token refresh errors: check tenant policies and cached token store permissions.
- Attachment download errors: ensure your access token has the Files/Attachments permission required to access embedded or external attachments.

## Contributing
Contributions and issues are welcome. Please:
1. Open an issue describing the bug or feature.
2. Fork the repo, create a branch for your changes.
3. Submit a pull request with a clear description and test steps.

Follow standard coding style and include tests where applicable.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Acknowledgments
- Microsoft Graph documentation and MSAL libraries
- Tkinter and the Python standard library for GUI and core features
