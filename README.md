# Certificate Project

A two-stage Python tool to generate branded PDF certificates and email them to webinar attendees via Microsoft 365.

Ensure .env files are loaded before running.
---

## Project Structure

```
certificate_project/
│
├── generate_certificates.py   # Stage 1 — generate PDFs
├── send_certificates.py       # Stage 2 — email attendees
├── requirements.txt           # Python dependencies
├── attendees.xlsx             # Your attendee list (you create this)
│
├── assets/
│   ├── logo.png               # Add your business logo here
│   └── medal.png              # Optional — add your own medal image
│
└── certificates/              # Auto-created — generated PDFs appear here
```

---

## Setup

### 1. Install dependencies
```bash
pip install reportlab openpyxl pillow
```

### 2. Prepare your attendees.xlsx
Create an Excel file with these exact column headers in row 1:

| email | full name | course name | date |
|-------|---------------|--------------------|----|
| jane.smith@email.com | Jane Smith | Dummy Leadership Course | 2026-03-22 |
| john.doe@email.com | John Doe | Dummy Leadership Course | 2026-03-22 |

### 3. Add your logo
Place your logo file at: `assets/logo.png`

### 4. Update configuration
In `generate_certificates.py`, edit the CONFIGURATION section:
```python
BUSINESS_NAME = "Your Actual Business Name"
BRAND_COLOUR = colors.HexColor("#YOUR_HEX")
ACCENT_COLOUR = colors.HexColor("#YOUR_HEX")
```

In .env, edit:
```python
SENDER_EMAIL = "your@actual365email.com"
BUSINESS_NAME = "Your Actual Business Name"
```

---

## Running the Project

### Stage 1 — Generate certificates
```bash
python generate_certificates.py
```
Check the `/certificates` folder — one PDF per attendee.

### Stage 2 — Email certificates
```bash
python send_certificates.py
```
You will be prompted for your M365 password at runtime. It is never stored.

---

## Microsoft 365 Notes

- SMTP host: `smtp.office365.com`, port `587`
- Uses Azure Graphs to send email.

---

## Security Rules

- Never hardcode passwords in the scripts
- Never commit `attendees.xlsx` to a public git repo — it contains personal data
- Add `attendees.xlsx` and `certificates/` to your `.gitignore`

---

## Customisation

| What | Where |
|------|-------|
| Brand colours | `generate_certificates.py` CONFIGURATION section |
| Email body text | `send_certificates.py` EMAIL_BODY_TEMPLATE |
| Certificate layout | `generate_certificates.py` generate_certificate() function |
| Medal image | Set `MEDAL_PATH = "assets/medal.png"` in generate_certificates.py |
