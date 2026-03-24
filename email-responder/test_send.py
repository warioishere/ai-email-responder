#!/usr/bin/env python3
"""Test email sending via SMTP + IMAP Sent folder save."""
import smtplib
import imaplib
import time
import yaml
from email.message import EmailMessage

with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f)

to = input("An (email): ").strip() or config['email']
subject = input("Betreff [Test]: ").strip() or "Test"
body = input("Text [Dies ist ein Test.]: ").strip() or "Dies ist ein Test."

msg = EmailMessage()
msg['From'] = config['email']
msg['To'] = to
msg['Subject'] = subject
msg.set_content(body)

# SMTP senden
print(f"\nSende via SMTP ({config['smtp_server']}:{config['smtp_port']})...")
try:
    with smtplib.SMTP(config['smtp_server'], config['smtp_port'], timeout=30) as smtp:
        smtp.starttls()
        smtp.login(config['email'], config['password'])
        smtp.send_message(msg)
    print(f"✓ Gesendet an {to}")
except Exception as e:
    print(f"✗ SMTP Fehler: {e}")
    exit(1)

# IMAP Sent folder
print("\nSpeichere in IMAP Sent folder...")
try:
    imap = imaplib.IMAP4_SSL(config['imap_server'])
    imap.login(config['email'], config['password'])
    sent_raw = msg.as_bytes()
    sent_time = imaplib.Time2Internaldate(time.time())
    for folder in ('INBOX.Sent', 'Sent', 'Sent Messages'):
        try:
            result = imap.append(folder, '\\Seen', sent_time, sent_raw)
            if result[0] == 'OK':
                print(f"✓ Gespeichert in: {folder}")
                break
        except Exception:
            continue
    else:
        print("✗ Kein Sent-Folder gefunden")
    imap.logout()
except Exception as e:
    print(f"✗ IMAP Fehler: {e}")

print("\nDone.")
