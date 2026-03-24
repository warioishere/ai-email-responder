import imaplib
import smtplib
import email
from email.header import decode_header
import os
import socket
import sqlite3
from email.message import EmailMessage
from anthropic import Anthropic
from datetime import datetime
import json
import time
from typing import List, Dict, Optional
import yaml
import re
from datetime import timedelta
import httpx
import uuid
import threading

class EmailAssistant:
    def __init__(self, config_path: str = 'config.yaml'):
        """Initialize the email assistant with configuration."""
        self.load_config(config_path)
        self.anthropic = Anthropic(api_key=self.config['anthropic_api_key'])
        self._init_db()
        self.connect_imap()
        self._load_article_index()

    def load_config(self, config_path: str):
        """Load configuration from YAML file."""
        with open(config_path, 'r') as f:
            self.config = yaml.safe_load(f)

    def _init_db(self):
        """Initialize SQLite database with all tables."""
        os.makedirs('memory', exist_ok=True)
        self.db_path = os.path.join('memory', 'email_assistant.db')
        self._db_lock = threading.Lock()
        self._db = sqlite3.connect(self.db_path, check_same_thread=False)
        self._db.execute("PRAGMA journal_mode=WAL")
        self._db.execute("PRAGMA busy_timeout=5000")
        self._db.row_factory = sqlite3.Row

        self._db.executescript("""
            CREATE TABLE IF NOT EXISTS pending (
                id TEXT PRIMARY KEY,
                type TEXT NOT NULL DEFAULT 'decision',
                sender TEXT NOT NULL,
                subject TEXT NOT NULL DEFAULT '',
                content TEXT DEFAULT '',
                message_id TEXT DEFAULT '',
                triage_category TEXT DEFAULT '',
                triage_confidence REAL DEFAULT 0,
                triage_reason TEXT DEFAULT '',
                draft TEXT,
                draft_raw TEXT,
                appointment_stage TEXT,
                appointment_time TEXT,
                resolved INTEGER NOT NULL DEFAULT 0,
                created TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS contacts (
                email TEXT PRIMARY KEY,
                name TEXT DEFAULT '',
                category_tags TEXT DEFAULT '[]',
                topics TEXT DEFAULT '[]',
                first_contact TEXT,
                last_contact TEXT
            );

            CREATE TABLE IF NOT EXISTS conversations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sender TEXT NOT NULL,
                subject TEXT DEFAULT '',
                email_content TEXT DEFAULT '',
                response TEXT DEFAULT '',
                created TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_conv_sender ON conversations(sender);
            CREATE INDEX IF NOT EXISTS idx_conv_created ON conversations(created);

            CREATE TABLE IF NOT EXISTS spam_senders (
                sender TEXT PRIMARY KEY
            );

            CREATE TABLE IF NOT EXISTS spam_keywords (
                keyword TEXT PRIMARY KEY
            );

            CREATE TABLE IF NOT EXISTS processed_ids (
                message_id TEXT PRIMARY KEY,
                type TEXT NOT NULL,
                created TEXT
            );

            CREATE TABLE IF NOT EXISTS pending_drafts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                recipient TEXT NOT NULL,
                subject TEXT NOT NULL DEFAULT '',
                original_content TEXT DEFAULT '',
                draft_response TEXT DEFAULT '',
                original_message_id TEXT DEFAULT '',
                calendar_appointment TEXT,
                created TEXT NOT NULL
            );
        """)
        self._db.commit()

        # Migrate old JSON files if they exist
        self._migrate_json_to_db()

        # Cache spam data for fast blacklist checks
        self._spam_senders = set(r[0] for r in self._db.execute("SELECT sender FROM spam_senders").fetchall())
        self._spam_keywords = [r[0] for r in self._db.execute("SELECT keyword FROM spam_keywords").fetchall()]
        print(f"✓ Database initialized ({self.db_path})")
        print(f"  Spam senders: {len(self._spam_senders)}, keywords: {len(self._spam_keywords)}")

    def _migrate_json_to_db(self):
        """Migrate existing JSON files to SQLite, then rename to .bak."""
        migrated = []

        # Migrate learned_spam.json
        if os.path.exists('learned_spam.json'):
            try:
                with open('learned_spam.json', 'r') as f:
                    spam = json.load(f)
                with self._db_lock:
                    for s in spam.get('senders', []):
                        self._db.execute("INSERT OR IGNORE INTO spam_senders(sender) VALUES(?)", (s,))
                    for k in spam.get('keywords', []):
                        self._db.execute("INSERT OR IGNORE INTO spam_keywords(keyword) VALUES(?)", (k,))
                    for mid in spam.get('processed_message_ids', []):
                        self._db.execute("INSERT OR IGNORE INTO processed_ids(message_id, type, created) VALUES(?, 'spam_learned', ?)",
                                        (mid, datetime.now().isoformat()))
                    self._db.commit()
                os.rename('learned_spam.json', 'learned_spam.json.bak')
                migrated.append('learned_spam.json')
            except Exception as e:
                print(f"  Error migrating learned_spam.json: {e}")

        # Migrate draft_tracking.json
        if os.path.exists('draft_tracking.json'):
            try:
                with open('draft_tracking.json', 'r') as f:
                    tracking = json.load(f)
                with self._db_lock:
                    for mid in tracking.get('processed_incoming_ids', []):
                        self._db.execute("INSERT OR IGNORE INTO processed_ids(message_id, type, created) VALUES(?, 'incoming', ?)",
                                        (mid, datetime.now().isoformat()))
                    for mid in tracking.get('learned_from', []):
                        self._db.execute("INSERT OR IGNORE INTO processed_ids(message_id, type, created) VALUES(?, 'learned', ?)",
                                        (mid, datetime.now().isoformat()))
                    for mid in tracking.get('manually_sent_learned', []):
                        self._db.execute("INSERT OR IGNORE INTO processed_ids(message_id, type, created) VALUES(?, 'manually_learned', ?)",
                                        (mid, datetime.now().isoformat()))
                    for draft in tracking.get('pending_drafts', []):
                        cal = json.dumps(draft.get('calendar_appointment')) if draft.get('calendar_appointment') else None
                        self._db.execute("""INSERT INTO pending_drafts(recipient, subject, original_content, draft_response,
                                           original_message_id, calendar_appointment, created)
                                           VALUES(?,?,?,?,?,?,?)""",
                                        (draft['recipient'], draft['subject'], draft.get('original_content', ''),
                                         draft.get('draft_response', ''), draft.get('original_message_id', ''),
                                         cal, draft.get('timestamp', datetime.now().isoformat())))
                    self._db.commit()
                os.rename('draft_tracking.json', 'draft_tracking.json.bak')
                migrated.append('draft_tracking.json')
            except Exception as e:
                print(f"  Error migrating draft_tracking.json: {e}")

        # Migrate pending_decisions.json
        pending_path = os.path.join('memory', 'categories', 'pending_decisions.json')
        if os.path.exists(pending_path):
            try:
                with open(pending_path, 'r') as f:
                    data = json.load(f)
                items = data.values() if isinstance(data, dict) else data
                with self._db_lock:
                    for item in items:
                        pid = item.get('id') or uuid.uuid4().hex[:6]
                        self._db.execute("""INSERT OR IGNORE INTO pending(id, type, sender, subject, content, message_id,
                                           triage_category, triage_confidence, triage_reason, draft, draft_raw,
                                           appointment_stage, appointment_time, resolved, created)
                                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                                        (pid, item.get('type', 'decision'), item.get('sender', ''),
                                         item.get('subject', ''), item.get('content', ''),
                                         item.get('message_id', ''), item.get('triage_category', ''),
                                         item.get('triage_confidence', 0), item.get('triage_reason', ''),
                                         item.get('draft'), item.get('draft_raw'),
                                         item.get('appointment_stage'), item.get('appointment_time'),
                                         1 if item.get('resolved') else 0,
                                         item.get('created', datetime.now().isoformat())))
                    self._db.commit()
                os.rename(pending_path, pending_path + '.bak')
                migrated.append('pending_decisions.json')
            except Exception as e:
                print(f"  Error migrating pending_decisions.json: {e}")

        # Migrate per-contact conversation files
        contacts_dir = os.path.join('memory', 'contacts')
        if os.path.isdir(contacts_dir):
            contact_files = [f for f in os.listdir(contacts_dir) if f.endswith('.json')]
            if contact_files:
                try:
                    with self._db_lock:
                        for filename in contact_files:
                            filepath = os.path.join(contacts_dir, filename)
                            with open(filepath, 'r') as f:
                                contact = json.load(f)
                            email_addr = contact.get('email', '')
                            if not email_addr:
                                continue
                            self._db.execute("""INSERT OR IGNORE INTO contacts(email, name, category_tags, topics, first_contact, last_contact)
                                               VALUES(?,?,?,?,?,?)""",
                                            (email_addr, contact.get('name', ''),
                                             json.dumps(contact.get('category_tags', [])),
                                             json.dumps(contact.get('topics', [])),
                                             contact.get('first_contact', ''),
                                             contact.get('last_contact', '')))
                            for conv in contact.get('conversations', []):
                                self._db.execute("""INSERT INTO conversations(sender, subject, email_content, response, created)
                                                   VALUES(?,?,?,?,?)""",
                                                (email_addr, conv.get('subject', ''),
                                                 conv.get('email_content', ''), conv.get('response', ''),
                                                 conv.get('date', datetime.now().isoformat())))
                        self._db.commit()
                    # Rename contact dir
                    os.rename(contacts_dir, contacts_dir + '.bak')
                    migrated.append(f'{len(contact_files)} contact files')
                except Exception as e:
                    print(f"  Error migrating contacts: {e}")

        # Migrate flat conversation_history.json (oldest format)
        if os.path.exists('conversation_history.json'):
            try:
                with open('conversation_history.json', 'r') as f:
                    old_history = json.load(f)
                with self._db_lock:
                    for email_addr, conversations in old_history.items():
                        self._db.execute("""INSERT OR IGNORE INTO contacts(email, name, first_contact, last_contact)
                                           VALUES(?,?,?,?)""",
                                        (email_addr, email_addr.split('@')[0].replace('.', ' ').title(),
                                         conversations[0]['date'] if conversations else '',
                                         conversations[-1]['date'] if conversations else ''))
                        for conv in conversations:
                            self._db.execute("""INSERT INTO conversations(sender, subject, email_content, response, created)
                                               VALUES(?,?,?,?,?)""",
                                            (email_addr, conv.get('subject', ''),
                                             conv.get('email_content', ''), conv.get('response', ''),
                                             conv.get('date', '')))
                    self._db.commit()
                os.rename('conversation_history.json', 'conversation_history.json.bak')
                migrated.append('conversation_history.json')
            except Exception as e:
                print(f"  Error migrating conversation_history.json: {e}")

        if migrated:
            print(f"  Migrated to DB: {', '.join(migrated)}")

    # ---- DB helpers ----

    def _db_is_processed(self, message_id: str, msg_type: str) -> bool:
        """Check if a message ID has been processed."""
        row = self._db.execute("SELECT 1 FROM processed_ids WHERE message_id=? AND type=?",
                              (message_id, msg_type)).fetchone()
        return row is not None

    def _db_mark_processed(self, message_id: str, msg_type: str):
        """Mark a message ID as processed."""
        with self._db_lock:
            self._db.execute("INSERT OR IGNORE INTO processed_ids(message_id, type, created) VALUES(?,?,?)",
                            (message_id, msg_type, datetime.now().isoformat()))
            self._db.commit()

    def _db_has_conversation(self, sender: str) -> bool:
        """Check if we have any conversation history with a sender."""
        row = self._db.execute("SELECT 1 FROM conversations WHERE sender=? LIMIT 1", (sender,)).fetchone()
        return row is not None

    def _db_get_pending(self, pid: str) -> Optional[Dict]:
        """Get a single pending item by ID."""
        row = self._db.execute("SELECT * FROM pending WHERE id=?", (pid,)).fetchone()
        return dict(row) if row else None

    def _db_update_pending(self, pid: str, **kwargs):
        """Update columns on a pending item."""
        if not kwargs:
            return
        cols = ', '.join(f"{k}=?" for k in kwargs)
        vals = list(kwargs.values()) + [pid]
        with self._db_lock:
            self._db.execute(f"UPDATE pending SET {cols} WHERE id=?", vals)
            self._db.commit()

    def mark_as_spam(self, email_data: Dict):
        """Mark an email as spam and learn keywords from it."""
        message_id = email_data.get('message_id', '')
        if message_id and self._db_is_processed(message_id, 'spam_learned'):
            return

        sender = email_data['sender']
        with self._db_lock:
            self._db.execute("INSERT OR IGNORE INTO spam_senders(sender) VALUES(?)", (sender,))
            self._spam_senders.add(sender)
            print(f"Added {sender} to spam sender list")

            # Extract 3-word phrases from subject
            subject_words = email_data['subject'].split()
            for i in range(len(subject_words)):
                if i + 2 < len(subject_words):
                    phrase = ' '.join(subject_words[i:i+3])
                    if len(phrase) > 10:
                        self._db.execute("INSERT OR IGNORE INTO spam_keywords(keyword) VALUES(?)", (phrase,))
                        if phrase not in self._spam_keywords:
                            self._spam_keywords.append(phrase)
                            print(f"Learned spam keyword: {phrase}")

            if message_id:
                self._db.execute("INSERT OR IGNORE INTO processed_ids(message_id, type, created) VALUES(?, 'spam_learned', ?)",
                                (message_id, datetime.now().isoformat()))
            self._db.commit()
        print(f"Email from {sender} marked as spam and patterns learned")

    def learn_from_spam_folder(self):
        """Check Junk/Spam folder and learn from emails there."""
        spam_folders = ['Junk', 'Spam', 'INBOX.Junk', 'INBOX.Spam']
        learned_count = 0

        # Calculate yesterday's date for filtering
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%d-%b-%Y')

        for folder in spam_folders:
            try:
                # Try to select the spam folder
                status, _ = self.imap.select(folder, readonly=True)
                if status != 'OK':
                    continue  # Folder doesn't exist, try next one

                print(f"Checking {folder} folder for spam emails since {yesterday}...")

                # Get emails in spam folder since yesterday
                _, message_numbers = self.imap.search(None, f'SINCE {yesterday}')

                if not message_numbers[0]:
                    continue  # No emails in this folder

                for num in message_numbers[0].split():
                    try:
                        _, msg_data = self.imap.fetch(num, '(RFC822)')
                        email_body = msg_data[0][1]
                        email_message = email.message_from_bytes(email_body)

                        content = self._extract_text_content(email_message)

                        sender = email.utils.parseaddr(email_message['From'])[1]
                        raw_subject = (email_message['Subject'] or '').replace('\n', ' ').replace('\r', ' ').strip()
                        subject = self.decode_mime_header(raw_subject)
                        message_id = (email_message['Message-ID'] or '').replace('\n', '').replace('\r', '').strip()

                        # Check if we've already learned from this email
                        if message_id and self._db_is_processed(message_id, 'spam_learned'):
                            continue

                        # Learn from this spam email
                        email_data = {
                            'sender': sender,
                            'subject': subject,
                            'content': content,
                            'message_id': message_id
                        }
                        self.mark_as_spam(email_data)
                        learned_count += 1

                    except Exception as e:
                        print(f"Error processing spam email: {str(e)}")
                        continue

                # Successfully processed this folder
                print(f"Learned from {learned_count} new spam emails in {folder}")
                return learned_count

            except Exception as e:
                # This folder doesn't exist or error accessing it
                continue

        if learned_count == 0:
            print("No new spam emails to learn from")

        return learned_count

    def connect_imap(self):
        """Connect to IMAP server."""
        try:
            self.imap = imaplib.IMAP4_SSL(self.config['imap_server'])
            self.imap.socket().settimeout(30)
            self.imap.login(self.config['email'], self.config['password'])
            print("✓ IMAP connection established")
        except Exception as e:
            print(f"✗ IMAP connection failed: {e}")
            raise

    def reconnect_imap(self):
        """Reconnect to IMAP server after connection loss."""
        try:
            print("Reconnecting to IMAP server...")
            try:
                self.imap.logout()
            except:
                pass  # Connection already dead

            self.imap = imaplib.IMAP4_SSL(self.config['imap_server'])
            self.imap.socket().settimeout(30)
            self.imap.login(self.config['email'], self.config['password'])
            print("✓ IMAP reconnection successful")
            return True
        except Exception as e:
            print(f"✗ IMAP reconnection failed: {e}")
            return False

    def _ensure_contact(self, email_addr: str):
        """Create contact record if it doesn't exist."""
        name = email_addr.split('@')[0].replace('.', ' ').replace('_', ' ').title()
        now = datetime.now().isoformat()
        with self._db_lock:
            self._db.execute("""INSERT OR IGNORE INTO contacts(email, name, first_contact, last_contact)
                               VALUES(?,?,?,?)""", (email_addr, name, now, now))
            self._db.commit()

    def _get_relevant_history(self, sender: str) -> str:
        """Get conversation history for a specific sender.
        Also cleans up conversations older than 14 weeks."""
        self._cleanup_old_conversations(sender, weeks=14)

        rows = self._db.execute(
            "SELECT subject, email_content, response, created FROM conversations WHERE sender=? ORDER BY created DESC LIMIT 3",
            (sender,)).fetchall()

        if not rows:
            return "No previous conversations with this sender."

        history_text = ""
        for row in reversed(rows):  # oldest first
            history_text += f"\nDate: {row['created']}\n"
            history_text += f"Email: {row['email_content']}\n"
            history_text += f"Response: {row['response']}\n"

        return history_text

    def _get_recent_emails_context(self, days: int = 10) -> str:
        """Build a summary of all conversations from the last N days for global context."""
        cutoff = (datetime.now() - timedelta(days=days)).isoformat()
        rows = self._db.execute(
            "SELECT sender, subject, email_content, response, created FROM conversations WHERE created>? ORDER BY created DESC LIMIT 20",
            (cutoff,)).fetchall()

        if not rows:
            return ""

        context = f"Recent email activity (last {days} days, {len(rows)} emails):\n"
        for row in rows:
            status = "replied" if row['response'] else "unread"
            context += f"- {row['created'][:10]} | {row['sender']} | {(row['subject'] or '')[:60]} [{status}]\n"

        return context

    def _cleanup_old_conversations(self, sender: str, weeks: int = 14):
        """Remove conversations older than the specified number of weeks."""
        cutoff = (datetime.now() - timedelta(weeks=weeks)).isoformat()
        with self._db_lock:
            result = self._db.execute("DELETE FROM conversations WHERE sender=? AND created<?", (sender, cutoff))
            if result.rowcount > 0:
                print(f"Cleaned up {result.rowcount} old conversations for {sender}")
                self._db.commit()

    def update_history(self, email_data: Dict, response: str):
        """Update conversation history with new email and response."""
        sender = email_data['sender']
        self._ensure_contact(sender)
        now = datetime.now().isoformat()
        with self._db_lock:
            self._db.execute("INSERT INTO conversations(sender, subject, email_content, response, created) VALUES(?,?,?,?,?)",
                            (sender, email_data.get('subject', ''), email_data.get('content', ''), response, now))
            self._db.execute("UPDATE contacts SET last_contact=? WHERE email=?", (now, sender))
            self._db.commit()

    def is_blacklisted(self, sender: str, subject: str, content: str) -> bool:
        """Check if email should be filtered out based on blacklist, automated senders, and keywords."""

        # Emails from own address: distinguish contact form vs order notifications
        own_email = self.config.get('email', '').lower()
        if own_email and sender.lower() == own_email:
            # Order notifications from WooCommerce -> filter (don't auto-reply)
            order_keywords = self.config.get('order_keywords', [])
            combined = f"{subject} {content}".lower()
            is_order = any(kw.lower() in combined for kw in order_keywords)
            if is_order:
                print(f"Email from own address {sender} is order notification, filtering")
                return True
            # Contact form submissions -> don't filter, let triage handle it
            print(f"Email from own address {sender} (contact form), not filtering")
            return False

        # FIRST: Check for automated/system senders (catches no-reply, MAILER-DAEMON, etc)
        # These should NEVER get responses
        automated_sender_patterns = [
            'no-reply@',
            'noreply@',
            'MAILER-DAEMON@',
            'postmaster@',
            'notification@',
            'alert@',
            'do-not-reply@',
            'automated@',
            'system@',
            'bounce@',
            'admin@',
            'robot@',
            'service@',
        ]

        sender_lower = sender.lower()
        for pattern in automated_sender_patterns:
            if pattern.lower() in sender_lower:
                print(f"Email from {sender} filtered out (automated sender: {pattern})")
                return True

        # Check sender blacklist
        blacklist = self.config.get('blacklist', [])
        for pattern in blacklist:
            if pattern.lower() in sender_lower:
                print(f"Email from {sender} is blacklisted (sender pattern: {pattern})")
                return True

        # Check learned spam senders
        if sender in self._spam_senders:
            print(f"Email from {sender} filtered out (learned spam sender)")
            return True

        # Check for order-related keywords
        order_keywords = self.config.get('order_keywords', [])
        combined_text = f"{subject} {content}".lower()
        for keyword in order_keywords:
            if keyword.lower() in combined_text:
                print(f"Email from {sender} filtered out (order keyword: {keyword})")
                return True

        # Check for advertising keywords
        ad_keywords = self.config.get('ad_keywords', [])
        for keyword in ad_keywords:
            if keyword.lower() in combined_text:
                print(f"Email from {sender} filtered out (ad keyword: {keyword})")
                return True

        # Check learned spam keywords
        for keyword in self._spam_keywords:
            if keyword.lower() in combined_text:
                print(f"Email from {sender} filtered out (learned spam keyword: {keyword})")
                return True

        return False

    def get_new_emails(self, search_criteria: str = 'UNSEEN') -> List[Dict]:
        """Fetch new emails from inbox."""
        self.imap.select('INBOX')

        # Calculate yesterday's date for filtering
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%d-%b-%Y')

        # Combine search criteria with date filter
        date_filter = f'SINCE {yesterday}'
        if search_criteria:
            combined_criteria = f'({search_criteria} {date_filter})'
            _, message_numbers = self.imap.search(None, search_criteria, date_filter)
        else:
            _, message_numbers = self.imap.search(None, date_filter)

        emails = []
        for num in message_numbers[0].split():
            _, msg_data = self.imap.fetch(num, '(BODY.PEEK[])')
            email_body = msg_data[0][1]
            email_message = email.message_from_bytes(email_body)

            content = self._extract_text_content(email_message)

            sender = email.utils.parseaddr(email_message['From'])[1]
            raw_subject = (email_message['Subject'] or '').replace('\n', ' ').replace('\r', ' ').strip()
            subject = self.decode_mime_header(raw_subject)
            message_id = (email_message['Message-ID'] or '').replace('\n', '').replace('\r', '').strip()

            # Check blacklist and filters
            # BUT: Don't filter if there's already a conversation with this sender
            # (e.g., customer replying to our manual response to their order confirmation)
            has_conversation = self._db_has_conversation(sender)

            if self.is_blacklisted(sender, subject, content) and not has_conversation:
                # Check what type of blocked email this is
                combined_text = f"{subject} {content}".lower()

                # Order confirmations: keep in inbox (no draft)
                order_keywords = self.config.get('order_keywords', [])
                is_order = any(kw.lower() in combined_text for kw in order_keywords)

                # DELETE everything EXCEPT order confirmations
                if not is_order:
                    # Move spam to Junk folder
                    print(f"Moving spam from {sender} to Junk")
                    self.move_to_junk(num)
                    continue
                else:
                    # Only keep order confirmations in inbox (no draft generated)
                    print(f"Filtering (keeping in inbox) order confirmation from {sender}")
                    continue

            elif has_conversation:
                print(f"Not filtering email from {sender} - existing conversation found")

            emails.append({
                'uid': num,
                'sender': sender,
                'subject': subject,
                'content': content,
                'message_id': message_id
            })

        return emails

    def mark_as_read(self, uid):
        """Mark email as read."""
        self.imap.store(uid, '+FLAGS', '\\Seen')

    def move_to_junk(self, uid):
        """Move email to Junk folder instead of deleting."""
        try:
            self.imap.select('INBOX')
            for folder in ['Junk', 'INBOX.Junk', 'Spam', 'INBOX.Spam']:
                try:
                    result = self.imap.copy(uid, folder)
                    if result[0] == 'OK':
                        self.imap.store(uid, '+FLAGS', '\\Deleted')
                        self.imap.expunge()
                        print(f"  Moved email {uid.decode() if isinstance(uid, bytes) else uid} to {folder}")
                        return
                except:
                    continue
            # Fallback: delete if no junk folder found
            print(f"  No Junk folder found, deleting instead")
            self.delete_email(uid)
        except Exception as e:
            print(f"  Error moving to junk: {e}")

    def delete_email(self, uid):
        """Permanently delete email from inbox."""
        try:
            self.imap.store(uid, '+FLAGS', '\\Deleted')
            self.imap.expunge()
            print(f"  Deleted email {uid.decode() if isinstance(uid, bytes) else uid}")
        except Exception as e:
            print(f"  Error deleting email {uid}: {e}")

    def _classify_email(self, email_data: Dict) -> Dict:
        """Classify an email into a triage category using a lightweight Claude call.
        Returns dict with category, confidence, and reason."""
        try:
            has_history = self._db_has_conversation(email_data['sender'])

            system_prompt = """You are an email triage classifier for a Swiss IT repair/privacy service business (yourdevice.ch).
Classify the incoming email into exactly ONE category:

- quick_answer: Simple questions answerable in under 5 minutes (product info, pricing, warranty, simple how-to, general questions, FAQ, thank-you notes, simple follow-ups)
- paid_consultation: Complex technical questions requiring hands-on help (GrapheneOS setup, Linux installation, hardware mods, data migration, networking, security/VPN/encryption setup, custom ROM, complex software config, plugin/module installation). These require a 135 CHF/h consultation.
- appointment_request: The sender wants to schedule a call, meeting, or visit. Or asks when you have time.
- needs_human: Too complex, personal, important, or ambiguous for auto-response. Complaints, legal matters, partnership proposals, anything requiring judgment or a personal decision.
- order_notification: Order confirmations, sale notifications, purchase receipts, shipping notifications from shops or marketplaces. Do NOT auto-reply to these unless there is an existing conversation (has_existing_conversation = true).
- spam: Unsolicited marketing, phishing, or irrelevant mass emails that got past filters.
- ignore: Automated notifications, delivery confirmations, read receipts, out-of-office replies.

IMPORTANT: If the sender is the same as the business email (info@yourdevice.ch), this is a CONTACT FORM submission from a potential customer - classify based on the actual question content, never as spam or ignore.

Respond with ONLY a JSON object, no other text: {"category": "...", "confidence": 0.XX, "reason": "..."}"""

            content_preview = email_data['content'][:500] if email_data['content'] else ''
            user_message = f"""Sender: {email_data['sender']}
Subject: {email_data['subject']}
Content: {content_preview}
Has existing conversation: {has_history}"""

            response = self.anthropic.messages.create(
                model=self.config['claude_model_name'],
                system=system_prompt,
                messages=[{"role": "user", "content": user_message}],
                max_tokens=150,
                temperature=0
            )

            result_text = response.content[0].text.strip()
            if not result_text:
                raise ValueError("Empty response from model")
            # Strip markdown code blocks if model wrapped the JSON
            if result_text.startswith('```'):
                result_text = re.sub(r'^```[a-z]*\n?', '', result_text)
                result_text = re.sub(r'\n?```$', '', result_text).strip()
            # Extract JSON object if there's surrounding text
            json_match = re.search(r'\{[^{}]+\}', result_text, re.DOTALL)
            if json_match:
                result_text = json_match.group(0)
            result = json.loads(result_text)

            valid_categories = ['quick_answer', 'paid_consultation', 'appointment_request',
                              'needs_human', 'order_notification', 'spam', 'ignore']
            if result.get('category') not in valid_categories:
                result['category'] = 'needs_human'
            if not isinstance(result.get('confidence'), (int, float)):
                result['confidence'] = 0.5

            return result

        except (json.JSONDecodeError, Exception) as e:
            print(f"Error classifying email: {e}")
            try:
                print(f"  Raw response was: {repr(result_text)}")
            except NameError:
                pass
            return {'category': 'needs_human', 'confidence': 0.0, 'reason': f'Classification error: {str(e)}'}

    def _save_pending_decision(self, email_data: Dict, triage: Dict) -> str:
        """Save email needing human decision. Returns the pending ID."""
        pid = uuid.uuid4().hex[:6]
        with self._db_lock:
            self._db.execute("""INSERT INTO pending(id, type, sender, subject, content, message_id,
                               triage_category, triage_confidence, triage_reason, created)
                               VALUES(?,?,?,?,?,?,?,?,?,?)""",
                            (pid, 'decision', email_data['sender'], email_data['subject'],
                             (email_data.get('content') or '')[:2000], email_data.get('message_id', ''),
                             triage['category'], triage['confidence'], triage.get('reason', ''),
                             datetime.now().isoformat()))
            self._db.commit()
        print(f"  Saved to pending [{pid}]: {triage['category']} - {triage.get('reason', '')}")
        return pid

    def _update_contact_from_triage(self, email_data: Dict, triage: Dict):
        """Update contact profile with triage information (category tags, topics)."""
        sender = email_data['sender']
        self._ensure_contact(sender)

        row = self._db.execute("SELECT category_tags, topics FROM contacts WHERE email=?", (sender,)).fetchone()
        if not row:
            return

        tags = json.loads(row['category_tags'] or '[]')
        topics = json.loads(row['topics'] or '[]')

        category = triage['category']
        if category in ('quick_answer', 'paid_consultation') and category not in tags:
            tags.append(category)

        reason = triage.get('reason', '')
        if reason and len(reason) > 3 and reason not in topics and len(topics) < 10:
            topics.append(reason)

        with self._db_lock:
            self._db.execute("UPDATE contacts SET category_tags=?, topics=?, last_contact=? WHERE email=?",
                            (json.dumps(tags), json.dumps(topics), datetime.now().isoformat(), sender))
            self._db.commit()

    # ---- Matrix Integration ----

    def _matrix_request(self, method: str, endpoint: str, json_data: dict = None) -> Optional[dict]:
        """Make an authenticated request to the Matrix client-server API."""
        homeserver = self.config.get('matrix_homeserver', '').rstrip('/')
        token = self.config.get('matrix_access_token', '')
        url = f"{homeserver}/_matrix/client/v3{endpoint}"
        headers = {"Authorization": f"Bearer {token}"}

        try:
            with httpx.Client(timeout=10) as client:
                if method == 'GET':
                    resp = client.get(url, headers=headers)
                elif method == 'PUT':
                    resp = client.put(url, headers=headers, json=json_data)
                elif method == 'POST':
                    resp = client.post(url, headers=headers, json=json_data)
                else:
                    return None

                if resp.status_code in (200, 201):
                    return resp.json()
                else:
                    print(f"Matrix API error {resp.status_code}: {resp.text[:200]}")
                    return None
        except Exception as e:
            print(f"Matrix request failed: {e}")
            return None

    def _matrix_send_html(self, html: str) -> bool:
        """Send a formatted HTML message. Plain text fallback is auto-generated."""
        # Strip HTML tags for plain text fallback
        plain = re.sub(r'<br\s*/?>', '\n', html)
        plain = re.sub(r'<hr\s*/?>', '---', plain)
        plain = re.sub(r'<[^>]+>', '', plain)
        plain = plain.strip()
        return self._matrix_send_message(plain, html)

    def _matrix_send_message(self, text: str, html: str = None) -> bool:
        """Send a message to the configured Matrix room."""
        if not self.config.get('matrix_enabled', False):
            return False

        room_id = self.config.get('matrix_room_id', '')
        txn_id = str(uuid.uuid4())
        endpoint = f"/rooms/{room_id}/send/m.room.message/{txn_id}"

        content = {"msgtype": "m.text", "body": text}
        if html:
            content["format"] = "org.matrix.custom.html"
            content["formatted_body"] = html

        result = self._matrix_request('PUT', endpoint, content)
        return result is not None

    def _matrix_notify_pending(self, email_data: Dict, triage: Dict, pending_id: str = None):
        """Send Matrix notification for email needing human review."""
        category = triage.get('category', 'unclassified')
        emoji = {"needs_human": "❓", "appointment_request": "📅", "quick_answer": "📧",
                 "paid_consultation": "💰", "order_notification": "📦"}.get(category, "📧")
        sender = email_data['sender']
        subject = email_data['subject']
        reason = triage.get('reason', '')
        content = (email_data.get('content') or '')[:600]
        pid = pending_id or '?'

        text = (f"{emoji} [{pid}] {category.upper()}\n"
                f"Von: {sender}\n"
                f"Betreff: {subject}\n"
                f"Grund: {reason}\n\n"
                f"Email:\n{content}\n\n"
                f"{pid} draft | {pid} draft [anweisung] | {pid} call | {pid} zeit [wann] | {pid} spam | {pid} ignore")

        html = (f"<b>{emoji} [{pid}] {category.upper()}</b><br/>"
                f"<b>Von:</b> {sender}<br/>"
                f"<b>Betreff:</b> {subject}<br/>"
                f"<b>Grund:</b> <i>{reason}</i><br/><br/>"
                f"<b>Email:</b><br/><pre>{content}</pre><br/>"
                f"<code>{pid} draft</code> | <code>{pid} draft [anweisung]</code> | <code>{pid} call</code> | "
                f"<code>{pid} zeit [wann]</code> | <code>{pid} spam</code> | <code>{pid} ignore</code>")

        if self._matrix_send_message(text, html):
            print(f"  Matrix notification sent for {sender} [{pid}]")
        else:
            print(f"  Failed to send Matrix notification for {sender}")

    def _matrix_notify_draft(self, email_data: Dict, draft: str, pending_id: str):
        """Send Matrix notification with auto-generated draft for confirmation."""
        sender = email_data['sender']
        subject = email_data['subject']
        content = (email_data.get('content') or '')[:400]
        cat = email_data.get('triage', {}).get('category', 'auto')
        pid = pending_id
        clean = re.sub(r'\s*CALENDAR_MARKER\|[^\n]*', '', draft).strip()

        text = (f"✅ [{pid}] DRAFT - {cat}\n"
                f"Von: {sender}\n"
                f"Betreff: {subject}\n\n"
                f"Email:\n{content}\n\n"
                f"Entwurf:\n{clean}\n\n"
                f"{pid} ok (senden) | {pid} ändern: [anweisung] | {pid} ignore")

        html = (f"<b>✅ [{pid}] DRAFT — {cat}</b><br/>"
                f"<b>Von:</b> {sender}<br/>"
                f"<b>Betreff:</b> {subject}<br/><br/>"
                f"<b>Email:</b><br/><pre>{content}</pre><br/>"
                f"<b>Entwurf:</b><br/><pre>{clean}</pre><br/>"
                f"<code>{pid} ok</code> (senden) | <code>{pid} ändern: [anweisung]</code> | <code>{pid} ignore</code>")

        if self._matrix_send_message(text, html):
            print(f"  Matrix draft notification sent [{pid}]")
        else:
            print(f"  Failed to send Matrix draft notification")

    def _matrix_loop(self):
        """Background thread: long-poll Matrix for messages and respond instantly."""
        print("Matrix listener started (long-polling)")
        sync_file = os.path.join('memory', 'matrix_sync_token.json')

        # Load saved sync token
        try:
            with open(sync_file, 'r') as f:
                self._matrix_since = json.load(f).get('next_batch')
        except (FileNotFoundError, json.JSONDecodeError):
            self._matrix_since = None

        # Initial sync (no timeout) to skip old messages
        if not self._matrix_since:
            result = self._matrix_sync(timeout=0)
            if result:
                self._matrix_since = result.get('next_batch')
                self._save_matrix_token()

        while True:
            try:
                self._matrix_check_responses()
            except Exception as e:
                print(f"Matrix loop error: {e}")
                time.sleep(5)

    def _matrix_sync(self, timeout: int = 30000) -> Optional[dict]:
        """Sync with Matrix server. timeout in milliseconds (30s default = long-poll)."""
        homeserver = self.config.get('matrix_homeserver', '').rstrip('/')
        token = self.config.get('matrix_access_token', '')
        room_id = self.config.get('matrix_room_id', '')

        url = f"{homeserver}/_matrix/client/v3/sync?timeout={timeout}"
        if self._matrix_since:
            url += f"&since={self._matrix_since}"
        url += f"&filter={{\"room\":{{\"rooms\":[\"{room_id}\"],\"timeline\":{{\"limit\":20}}}}}}"

        try:
            # Timeout slightly longer than the long-poll timeout
            with httpx.Client(timeout=max(timeout / 1000 + 10, 15)) as client:
                resp = client.get(url, headers={"Authorization": f"Bearer {token}"})
                if resp.status_code == 200:
                    return resp.json()
                else:
                    print(f"Matrix sync error {resp.status_code}")
                    return None
        except Exception as e:
            print(f"Matrix sync failed: {e}")
            return None

    def _save_matrix_token(self):
        """Save the Matrix sync token to disk."""
        sync_file = os.path.join('memory', 'matrix_sync_token.json')
        os.makedirs('memory', exist_ok=True)
        with open(sync_file, 'w') as f:
            json.dump({'next_batch': self._matrix_since}, f)

    def _matrix_check_responses(self):
        """Long-poll Matrix for new messages and process them."""
        if not self.config.get('matrix_enabled', False):
            return

        result = self._matrix_sync(timeout=30000)
        if not result:
            return

        next_batch = result.get('next_batch')
        if next_batch:
            self._matrix_since = next_batch
            self._save_matrix_token()

        room_id = self.config.get('matrix_room_id', '')
        rooms = result.get('rooms', {}).get('join', {})
        events = rooms.get(room_id, {}).get('timeline', {}).get('events', [])

        bot_user_id = None
        whoami = self._matrix_request('GET', '/account/whoami')
        if whoami:
            bot_user_id = whoami.get('user_id')

        for event in events:
            if event.get('type') != 'm.room.message':
                continue
            if event.get('sender') == bot_user_id:
                continue

            original_body = event.get('content', {}).get('body', '').strip()
            body = original_body.lower()
            if not body:
                continue

            if body == '!help':
                self._matrix_send_help()
                continue

            if body == '!status':
                self._matrix_send_status()
                continue

            # Parse: "<6-char-id> <command> [args]"
            parts = original_body.split(' ', 2)
            if len(parts) < 2:
                continue

            pid = parts[0].lower()
            command = parts[1].lower()
            args = parts[2] if len(parts) > 2 else ''

            item = self._db_get_pending(pid)
            if not item:
                self._matrix_send_message(f"❌ ID '{pid}' nicht gefunden. !status für aktuelle Liste.")
                continue

            if item['resolved']:
                self._matrix_send_message(f"⚠️ [{pid}] bereits erledigt.")
                continue

            if command == 'ok':
                self._mx_send(pid, item)
            elif command == 'draft':
                if args:
                    self._mx_regenerate(pid, item, args)
                else:
                    self._mx_generate_draft(pid, item, 'quick_answer')
            elif command == 'call':
                self._mx_generate_draft(pid, item, 'paid_consultation')
            elif command == 'zeit':
                self._mx_propose_appointment(pid, item, args)
            elif command == 'ändern':
                if not args:
                    self._matrix_send_message(f"❌ [{pid}] Was ändern? z.B. '{pid} ändern mach kürzer'")
                else:
                    self._mx_regenerate(pid, item, args)
            elif command == 'spam':
                self.mark_as_spam({'sender': item['sender'], 'subject': item['subject'],
                                   'content': item['content'] or '', 'message_id': item['message_id'] or ''})
                self._db_update_pending(pid, resolved=1)
                self._matrix_send_html(f"🗑️ <b>Spam [{pid}]:</b> {item['sender']}<br/><i>Wird ab jetzt gefiltert.</i>")
            elif command == 'ignore':
                self._db_update_pending(pid, resolved=1)
                self._matrix_send_html(f"🚫 <b>Ignoriert [{pid}]:</b> {item['sender']}")
            else:
                self._matrix_send_message(f"❓ Unbekannt: '{command}'. !help für Befehle.")

    def _matrix_send_help(self):
        open_count = self._db.execute("SELECT COUNT(*) FROM pending WHERE resolved=0").fetchone()[0]
        html = (f"<h4>📧 Email-Assistent</h4>"
                f"<b>Model:</b> <code>{self.config['claude_model_name']}</code><br/>"
                f"<b>Offene Items:</b> {open_count}<br/><br/>"
                f"<b>[id]</b> = 6-stellige ID aus der Notification<br/><br/>"
                f"<code>[id] draft</code> — Claude generiert Antwort<br/>"
                f"<code>[id] draft [anweisung]</code> — Claude generiert mit deiner Anweisung<br/>"
                f"<code>[id] call</code> — Beratungsangebot (135 CHF/h)<br/>"
                f"<code>[id] zeit [wann]</code> — Terminvorschlag erstellen<br/>"
                f"<code>[id] ok</code> — Draft senden<br/>"
                f"<code>[id] ändern [text]</code> — bestehenden Draft anpassen<br/>"
                f"<code>[id] spam</code> — Als Spam markieren<br/>"
                f"<code>[id] ignore</code> — Ignorieren<br/><br/>"
                f"<code>!status</code> — Alle offenen Items<br/>"
                f"<code>!help</code> — Diese Hilfe")
        self._matrix_send_html(html)

    def _matrix_send_status(self):
        rows = self._db.execute("SELECT * FROM pending WHERE resolved=0 ORDER BY created DESC").fetchall()
        if not rows:
            self._matrix_send_html("<i>Keine offenen Items.</i>")
            return
        html = f"<b>{len(rows)} offene Item(s):</b><br/><br/>"
        for row in rows:
            cat = row['triage_category'] or '?'
            has_draft = "✅ Draft bereit" if row['draft'] else "⏳ Kein Draft"
            html += (f"<b>[{row['id']}]</b> {row['sender']}<br/>"
                     f"&nbsp;&nbsp;📋 {row['subject']}<br/>"
                     f"&nbsp;&nbsp;<code>{cat}</code> | {has_draft}<br/><br/>")
        self._matrix_send_html(html)

    def _mx_generate_draft(self, pid: str, item: Dict, category: str):
        """Generate a draft for a pending item and show it in Matrix."""
        email_data = {
            'sender': item['sender'],
            'subject': item['subject'],
            'content': item['content'] or '',
            'message_id': item['message_id'] or '',
            'triage': {'category': category}
        }
        response = self.generate_response(email_data)
        if not response or response.startswith("Error"):
            self._matrix_send_message(f"❌ [{pid}] Fehler: {response}")
            return
        clean = re.sub(r'\s*CALENDAR_MARKER\|[^\n]*', '', response).strip()
        self._db_update_pending(pid, draft=clean, draft_raw=response)
        cat_label = "💰 Beratungsangebot (135 CHF/h)" if category == 'paid_consultation' else "✅ Draft"
        html = (f"<b>{cat_label} [{pid}]</b><br/>"
                f"<b>An:</b> {item['sender']}<br/><br/>"
                f"<b>Entwurf:</b><br/><pre>{clean}</pre><br/>"
                f"<code>{pid} ok</code> (senden) | <code>{pid} ändern [anweisung]</code>")
        self._matrix_send_html(html)

    def _mx_propose_appointment(self, pid: str, item: Dict, time_str: str):
        """Generate appointment proposal draft."""
        email_data = {
            'sender': item['sender'],
            'subject': item['subject'],
            'content': (item['content'] or '') + f"\n\n[SYSTEM: Schlage folgenden Termin vor: {time_str}]",
            'message_id': item['message_id'] or '',
            'triage': {'category': 'quick_answer'}
        }
        response = self.generate_response(email_data)
        if not response or response.startswith("Error"):
            self._matrix_send_message(f"❌ [{pid}] Fehler: {response}")
            return
        clean = re.sub(r'\s*CALENDAR_MARKER\|[^\n]*', '', response).strip()
        self._db_update_pending(pid, draft=clean, draft_raw=response, appointment_stage='proposed', appointment_time=time_str)
        html = (f"<b>📅 Terminvorschlag [{pid}]</b><br/>"
                f"<b>An:</b> {item['sender']}<br/>"
                f"<b>Zeit:</b> {time_str}<br/><br/>"
                f"<b>Entwurf:</b><br/><pre>{clean}</pre><br/>"
                f"<code>{pid} ok</code> (senden + Kalender) | <code>{pid} ändern [anweisung]</code>")
        self._matrix_send_html(html)

    def _mx_regenerate(self, pid: str, item: Dict, instructions: str):
        """Regenerate draft with custom instructions."""
        email_data = {
            'sender': item['sender'],
            'subject': item['subject'],
            'content': (item['content'] or '') + f"\n\n[SYSTEM: {instructions}]",
            'message_id': item['message_id'] or '',
            'triage': {'category': 'quick_answer'}
        }
        response = self.generate_response(email_data)
        if not response or response.startswith("Error"):
            self._matrix_send_message(f"❌ [{pid}] Fehler: {response}")
            return
        clean = re.sub(r'\s*CALENDAR_MARKER\|[^\n]*', '', response).strip()
        self._db_update_pending(pid, draft=clean, draft_raw=response)
        html = (f"<b>✏️ Draft aktualisiert [{pid}]</b><br/>"
                f"<b>An:</b> {item['sender']}<br/><br/>"
                f"<b>Entwurf:</b><br/><pre>{clean}</pre><br/>"
                f"<code>{pid} ok</code> (senden) | <code>{pid} ändern [anweisung]</code>")
        self._matrix_send_html(html)

    def _mx_send(self, pid: str, item: Dict):
        """Send the draft for a pending item via SMTP."""
        if not item['draft']:
            self._matrix_send_message(f"❌ [{pid}] Kein Draft. Erst '{pid} draft' aufrufen.")
            return
        to = item['sender']
        subject = item['subject'] if item['subject'].startswith('Re:') else f"Re: {item['subject']}"
        body_raw = item['draft_raw'] or item['draft']
        in_reply_to = item['message_id'] or ''

        if self.send_via_smtp(to, subject, body_raw, in_reply_to):
            email_data = {'sender': item['sender'], 'subject': item['subject'],
                          'content': item['content'] or '', 'message_id': item['message_id'] or ''}
            self.update_history(email_data, item['draft'])
            self._db_update_pending(pid, resolved=1)
            cal_note = ""
            if item['appointment_stage'] == 'proposed' and item['appointment_time']:
                cal_note = f"<br/>📅 Kalender-Eintrag für: {item['appointment_time']}"
            self._matrix_send_html(f"📤 <b>Gesendet [{pid}]</b><br/>An: {to}{cal_note}")
        else:
            self._matrix_send_message(f"❌ [{pid}] SMTP fehlgeschlagen. Manuell senden.")

    def _test_matrix_connection(self):
        """Test Matrix connection on startup."""
        if not self.config.get('matrix_enabled', False):
            print("Matrix integration: DISABLED")
            return False

        print("Testing Matrix connection...")
        result = self._matrix_request('GET', '/account/whoami')
        if result:
            user_id = result.get('user_id', 'unknown')
            print(f"Matrix connection OK: {user_id}")
            # Send startup message
            self._matrix_send_html("🟢 <b>Email-Assistent gestartet.</b> Schreib <code>!help</code> fuer Befehle.")
            return True
        else:
            print("Matrix connection FAILED")
            return False

    # ---- End Matrix Integration ----

    # ---- Article Index ----

    def _load_article_index(self):
        """Load the article index from memory/article_index.json"""
        index_path = os.path.join('memory', 'article_index.json')
        try:
            with open(index_path, 'r') as f:
                self._article_index = json.load(f)
            print(f"Article index loaded: {len(self._article_index)} articles")
        except (FileNotFoundError, json.JSONDecodeError):
            self._article_index = []
            print("No article index found (run article indexer to create one)")

    def _find_relevant_articles(self, email_data: Dict, max_results: int = 5) -> str:
        """Find relevant articles based on email subject and content using keyword matching
        against article titles, headings, and categories."""
        if not self._article_index or not self.config.get('article_linking_enabled', True):
            return ""

        subject = email_data.get('subject', '').lower()
        content = email_data.get('content', '')[:500].lower()
        search_text = f"{subject} {content}"
        search_words = set(re.findall(r'\w{4,}', search_text))

        if not search_words:
            return ""

        scored = []
        for article in self._article_index:
            score = 0

            # Title match (highest weight)
            title_words = set(re.findall(r'\w{4,}', article['title'].lower()))
            score += len(title_words & search_words) * 3

            # Headings match (good signal)
            for heading in article.get('headings', []):
                heading_words = set(re.findall(r'\w{4,}', heading.lower()))
                score += len(heading_words & search_words)

            # Category match
            cat_text = ' '.join(article.get('categories', [])).lower()
            cat_words = set(re.findall(r'\w{4,}', cat_text))
            score += len(cat_words & search_words)

            if score > 1:  # Minimum 2 keyword matches to avoid noise
                scored.append((score, article))

        if not scored:
            return ""

        # Sort by score, take top results
        scored.sort(key=lambda x: x[0], reverse=True)
        top = scored[:max_results]

        result = "Relevante Artikel von yourdevice.ch (verweise den Kunden darauf wenn passend):\n"
        for score, article in top:
            result += f"- {article['title']}: {article['url']}\n"

        return result

    # ---- End Article Index ----

    def generate_response(self, email_data: Dict) -> str:
        """Generate response using Claude API."""
        try:
            full_system_prompt = self.config.get('system_prompt', '')

            # Add triage-specific guidance
            triage = email_data.get('triage', {})
            category = triage.get('category', 'quick_answer')
            if category == 'paid_consultation':
                full_system_prompt += "\n\nWICHTIG: Diese Email erfordert eine kostenpflichtige Beratung. "
                full_system_prompt += "Antworte locker und erklaere dass diese Art von Hilfe eine bezahlte Session braucht (135 CHF/Stunde). "
                full_system_prompt += "Schlage vor einen Termin zu machen (persoenlich oder remote).\n"
            elif category == 'quick_answer':
                full_system_prompt += "\n\nEinfache Frage - antworte direkt und kurz.\n"

            # Current email context
            email_context = f"""
Von: {email_data['sender']}
Betreff: {email_data['subject']}
Inhalt: {email_data['content']}"""

            # Get conversation history for this sender
            history_context = self._get_relevant_history(email_data['sender'])

            # Get recent email activity for broader context
            recent_context = self._get_recent_emails_context(days=10)

            # Find relevant articles
            articles_context = self._find_relevant_articles(email_data)

            # Combine all context
            full_context = f"Bisherige Konversation mit diesem Absender:\n{history_context}"
            if recent_context:
                full_context += f"\n\n{recent_context}"
            if articles_context:
                full_context += f"\n\n{articles_context}"
            full_context += f"\n\nNeue Email beantworten:\n{email_context}"

            response = self.anthropic.messages.create(
                model=self.config['claude_model_name'],
                system=full_system_prompt,
                messages=[
                    {
                        "role": "user",
                        "content": full_context
                    }
                ],
                max_tokens=self.config.get('max_tokens', 1000),
                temperature=self.config.get('temperature', 0.7)
            )

            # Extract the response content safely
            if hasattr(response, 'content'):
                if isinstance(response.content, list):
                    # If content is a list, join all text parts
                    return ' '.join(item.text for item in response.content if hasattr(item, 'text'))
                return str(response.content)
            else:
                # Fallback for different response structure
                return str(response.messages[0].content)

        except Exception as e:
            print(f"Error generating response: {str(e)}")
            # Log the full error details
            import traceback
            print(f"Full error: {traceback.format_exc()}")
            return "Error generating response. Please check the logs for details."

    def decode_mime_header(self, header_value: str) -> str:
        """Decode MIME encoded header (e.g., =?UTF-8?Q?...?=) to plain text."""
        if not header_value:
            return ''
        try:
            decoded_parts = decode_header(header_value)
            result = ''
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    if encoding:
                        result += part.decode(encoding)
                    else:
                        result += part.decode('utf-8', errors='ignore')
                else:
                    result += str(part) if part else ''
            return result.strip()
        except:
            return header_value.strip()

    def _extract_text_content(self, email_message) -> str:
        """Extract plain text content from an email.message.Message object.
        Handles multipart messages, UTF-8 decoding with fallback encodings."""
        if email_message.is_multipart():
            content = ''
            for part in email_message.walk():
                if part.get_content_type() == "text/plain":
                    try:
                        payload = part.get_payload(decode=True)
                        if payload:
                            try:
                                content += payload.decode('utf-8')
                            except UnicodeDecodeError:
                                for encoding in ['iso-8859-1', 'windows-1252', 'latin-1']:
                                    try:
                                        content += payload.decode(encoding)
                                        break
                                    except UnicodeDecodeError:
                                        continue
                                else:
                                    content += payload.decode('utf-8', errors='ignore')
                    except:
                        pass
            return content
        else:
            try:
                payload = email_message.get_payload(decode=True)
                if payload:
                    try:
                        return payload.decode('utf-8')
                    except UnicodeDecodeError:
                        for encoding in ['iso-8859-1', 'windows-1252', 'latin-1']:
                            try:
                                return payload.decode(encoding)
                            except UnicodeDecodeError:
                                continue
                        return payload.decode('utf-8', errors='ignore')
                return ''
            except:
                return ''

    def remove_markdown(self, text: str) -> str:
        """Remove markdown formatting from text."""
        # Remove markdown bold (**text** -> text)
        text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
        # Remove markdown italic (*text* -> text)
        text = re.sub(r'\*(.*?)\*', r'\1', text)
        # Remove markdown headers (# text -> text)
        text = re.sub(r'^#+\s+', '', text, flags=re.MULTILINE)
        # Remove markdown bold underscores (__text__ -> text)
        text = re.sub(r'__(.*?)__', r'\1', text)
        # Remove markdown code blocks (```...``` -> ...)
        text = re.sub(r'```.*?```', '', text, flags=re.DOTALL)
        # Remove inline code (`text` -> text)
        text = re.sub(r'`(.*?)`', r'\1', text)
        return text

    def save_draft(self, email_data: Dict, response: str):
        """Save response as draft email."""
        # Clean subject and message_id of any newlines/carriage returns
        subject = email_data['subject'].replace('\n', ' ').replace('\r', ' ').strip()
        message_id = email_data['message_id'].replace('\n', '').replace('\r', '').strip()

        # Strip CALENDAR_MARKER from IMAP draft (kept in tracking for N send)
        clean_response = re.sub(r'\s*CALENDAR_MARKER\|[^\n]*', '', self.remove_markdown(response)).strip()

        draft = EmailMessage()
        draft['To'] = email_data['sender']
        draft['Subject'] = f"Re: {subject}"
        if message_id:
            draft['In-Reply-To'] = message_id
        draft.set_content(clean_response)

        # Save to drafts folder
        self.imap.append('Drafts', '', imaplib.Time2Internaldate(time.time()),
                        draft.as_bytes())

        # Parse calendar data now and store in tracking
        appointment = self.parse_calendar_marker(response, email_data['sender'], subject)
        cal_json = None
        if appointment:
            cal_json = json.dumps({
                'title': appointment['title'],
                'start': appointment['start'].isoformat(),
                'end': appointment['end'].isoformat(),
                'location': appointment.get('location', ''),
                'description': appointment.get('description', '')
            })

        with self._db_lock:
            self._db.execute("""INSERT INTO pending_drafts(recipient, subject, original_content, draft_response,
                               original_message_id, calendar_appointment, created) VALUES(?,?,?,?,?,?,?)""",
                            (email_data['sender'], email_data['subject'], email_data['content'],
                             response, email_data['message_id'], cal_json, datetime.now().isoformat()))
            self._db.commit()

    def send_via_smtp(self, to: str, subject: str, body: str, in_reply_to: str = '') -> bool:
        """Send an email via SMTP and remove the corresponding draft from IMAP."""
        try:
            # Parse and strip CALENDAR_MARKER before sending
            appointment = self.parse_calendar_marker(body, to, subject)
            clean_body = re.sub(r'\s*CALENDAR_MARKER\|[^\n]*', '', body).strip()

            msg = EmailMessage()
            msg['From'] = self.config['email']
            msg['To'] = to
            msg['Subject'] = subject
            if in_reply_to:
                msg['In-Reply-To'] = in_reply_to
                msg['References'] = in_reply_to
            msg.set_content(self.remove_markdown(clean_body))

            smtp_server = self.config.get('smtp_server', self.config['imap_server'])
            smtp_port = self.config.get('smtp_port', 587)

            with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as smtp:
                smtp.starttls()
                smtp.login(self.config['email'], self.config['password'])
                smtp.send_message(msg)

            print(f"✓ Email sent to {to}: {subject}")

            # Save copy to IMAP Sent folder
            try:
                sent_raw = msg.as_bytes()
                sent_time = imaplib.Time2Internaldate(time.time())
                for sent_folder in ('INBOX.Sent', 'Sent', 'Sent Messages', 'INBOX.Sent Messages'):
                    try:
                        result = self.imap.append(sent_folder, '\\Seen', sent_time, sent_raw)
                        if result[0] == 'OK':
                            print(f"  Saved to IMAP Sent folder: {sent_folder}")
                            break
                    except Exception:
                        continue
                else:
                    print("  Could not save to IMAP Sent folder (tried all variants)")
            except Exception as e:
                print(f"  Error saving to Sent folder: {e}")

            # Create calendar event if marker was found
            if appointment and self.config.get('enable_calendar', True):
                if self.create_calendar_event(appointment):
                    print(f"  Calendar event created: {appointment.get('title')}")

            # Remove matching draft from IMAP Drafts folder
            try:
                self.imap.select('Drafts')
                if in_reply_to:
                    _, nums = self.imap.search(None, f'HEADER In-Reply-To "{in_reply_to}"')
                else:
                    _, nums = self.imap.search(None, f'HEADER Subject "{subject}" TO "{to}"')
                if nums[0]:
                    for num in nums[0].split():
                        self.imap.store(num, '+FLAGS', '\\Deleted')
                    self.imap.expunge()
                    print(f"  Draft removed from IMAP Drafts")
            except Exception as e:
                print(f"  Could not remove draft from IMAP: {e}")

            return True
        except Exception as e:
            print(f"✗ SMTP send failed: {e}")
            return False

    def find_email_by_message_id(self, message_id: str) -> Optional[Dict]:
        """Find an email by its Message-ID in the INBOX."""
        try:
            self.imap.select('INBOX')
            # Search for the email by Message-ID
            _, result = self.imap.search(None, f'HEADER Message-ID "{message_id}"')

            if result[0]:
                num = result[0].split()[0]
                _, msg_data = self.imap.fetch(num, '(RFC822)')
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)

                content = self._extract_text_content(email_message)

                sender = email.utils.parseaddr(email_message['From'])[1]
                raw_subject = (email_message['Subject'] or '').replace('\n', ' ').replace('\r', ' ').strip()
                subject = self.decode_mime_header(raw_subject)

                return {
                    'sender': sender,
                    'subject': subject,
                    'content': content,
                    'message_id': message_id
                }
        except Exception as e:
            print(f"Error finding email by Message-ID: {e}")

        return None

    def find_email_by_recipient_subject(self, recipient: str, subject: str) -> Optional[Dict]:
        """Find original email by searching in INBOX for emails from recipient with matching subject.
        Used as fallback when Message-ID search fails. Works for both order confirmations and regular emails.
        Handles non-ASCII characters by extracting ASCII-only keywords."""
        try:
            import re

            # Strip "Re:" and "RE:" prefixes from subject for searching
            search_subject = subject
            search_subject = re.sub(r'^re:\s*', '', search_subject, flags=re.IGNORECASE)
            search_subject = search_subject.strip()

            print(f"    Searching in INBOX for emails from {recipient}")
            print(f"    Search subject: {search_subject}")

            # Try to search in INBOX first (where customer emails arrive)
            folders_to_check = ['INBOX', 'INBOX.Archive', 'Archive']

            for folder in folders_to_check:
                try:
                    self.imap.select(folder)

                    # Extract ASCII-only keywords from subject (to avoid IMAP encoding errors)
                    # IMAP can't handle non-ASCII characters in search terms
                    ascii_words = []
                    for word in search_subject.split():
                        try:
                            word.encode('ascii')
                            # Only include words that are >2 chars and ASCII
                            if len(word) > 2:
                                ascii_words.append(word)
                        except UnicodeEncodeError:
                            # Skip non-ASCII words
                            continue

                    print(f"    ASCII keywords: {ascii_words}")

                    email_ids = []

                    # Try searching by each ASCII keyword
                    for keyword in ascii_words[:3]:  # Try first 3 ASCII keywords
                        print(f"    Trying search by keyword: {keyword}")
                        try:
                            _, result = self.imap.search(None, f'SUBJECT "{keyword}"')
                            if result[0]:
                                email_ids = result[0].split()
                                print(f"    Found {len(email_ids)} emails with '{keyword}'")
                                break  # Found matches, stop searching
                        except Exception as e:
                            print(f"    Keyword search failed: {e}")
                            continue

                    # Check the most recent emails from this folder
                    if email_ids:
                        print(f"    Checking {min(10, len(email_ids))} potential matches in {folder}")
                        for email_id in email_ids[-10:]:  # Check last 10
                            try:
                                _, msg_data = self.imap.fetch(email_id, '(RFC822)')
                                email_body = msg_data[0][1]
                                email_message = email.message_from_bytes(email_body)

                                # Get sender (FROM field)
                                sender = email.utils.parseaddr(email_message['From'])[1]

                                # Check if this is from the right person
                                if sender.lower() == recipient.lower():
                                    # Decode the subject
                                    raw_subject = (email_message['Subject'] or '').replace('\n', ' ').replace('\r', ' ').strip()
                                    decoded_subject = self.decode_mime_header(raw_subject)

                                    print(f"    Found matching email from {sender}: {decoded_subject}")

                                    content = self._extract_text_content(email_message)

                                    return {
                                        'sender': sender,
                                        'subject': decoded_subject,
                                        'content': content,
                                        'message_id': email_message.get('Message-ID', '')
                                    }
                            except Exception as e:
                                continue
                except Exception as e:
                    # Failed to search this folder, try next one
                    print(f"    Error searching {folder}: {e}")
                    continue

            print(f"    Could not find original email from {recipient}")
            return None
        except Exception as e:
            print(f"Error finding email by recipient+subject: {e}")
            return None

    def learn_from_sent_emails(self):
        """Check Sent folder and learn from emails that match our drafts."""
        sent_folders = ['Sent', 'INBOX.Sent', '[Gmail]/Sent Mail', 'Sent Items']
        learned_count = 0

        # Clean up old pending drafts (older than 30 days)
        cutoff = (datetime.now() - timedelta(days=30)).isoformat()
        with self._db_lock:
            result = self._db.execute("DELETE FROM pending_drafts WHERE created<?", (cutoff,))
            if result.rowcount > 0:
                print(f"Cleaned up {result.rowcount} old pending drafts (>30 days)")
                self._db.commit()

        # Check sent emails from last 7 days (covers weekends and short breaks)
        week_ago = (datetime.now() - timedelta(days=7)).strftime('%d-%b-%Y')

        for folder in sent_folders:
            try:
                # CRITICAL: Close current folder to force folder switch
                try:
                    self.imap.close()
                except:
                    pass  # Might fail if no folder is currently selected

                # Try to select the sent folder
                status, response = self.imap.select(folder, readonly=True)
                if status != 'OK':
                    continue  # Folder doesn't exist, try next one

                # Verify which folder is actually selected by fetching last email
                print(f"Checking {folder} folder for sent emails since {week_ago}...")
                print(f"Folder response: {response}")

                # Fetch the LAST email to verify we're in the right folder
                _, all_msgs = self.imap.search(None, 'ALL')
                if all_msgs[0]:
                    last_id = all_msgs[0].split()[-1]
                    _, test_msg = self.imap.fetch(last_id, '(RFC822.HEADER)')
                    test_email = email.message_from_bytes(test_msg[0][1])
                    test_subj = self.decode_mime_header((test_email['Subject'] or '').replace('\n', ' ').strip())
                    test_to = self.decode_mime_header((test_email['To'] or '').replace('\n', ' ').strip())
                    print(f"Last email in this folder (ID {last_id.decode()}):")
                    print(f"  Subject: {test_subj[:80]}")
                    print(f"  To: {test_to[:80]}")

                # Get emails sent in last 7 days
                _, message_numbers = self.imap.search(None, f'SINCE {week_ago}')

                if not message_numbers[0]:
                    print(f"No sent emails found in {folder} since {week_ago}")
                    continue  # No emails in this folder

                sent_emails = message_numbers[0].split()
                pending_draft_count = self._db.execute("SELECT COUNT(*) FROM pending_drafts").fetchone()[0]
                print(f"Found {len(sent_emails)} sent emails in {folder} to check")
                print(f"Pending drafts to match: {pending_draft_count}")

                for idx, num in enumerate(sent_emails):
                    try:
                        print(f"  [{idx+1}/{len(sent_emails)}] Fetching email ID {num.decode()}...")
                        # Fetch only headers first to check Message-ID before loading full body
                        _, hdr_data = self.imap.fetch(num, '(RFC822.HEADER)')
                        if not hdr_data or not hdr_data[0] or len(hdr_data[0]) < 2:
                            print(f"    ERROR: Failed to fetch headers for #{num.decode()}")
                            continue
                        hdr_message = email.message_from_bytes(hdr_data[0][1])
                        message_id = (hdr_message['Message-ID'] or '').replace('\n', '').replace('\r', '').strip()

                        # Skip if already learned - no need to fetch full body
                        if message_id and (self._db_is_processed(message_id, 'learned') or self._db_is_processed(message_id, 'manually_learned')):
                            print(f"    Already learned from this sent email, skipping")
                            continue

                        # Fetch headers + text body only (no attachments)
                        _, msg_data = self.imap.fetch(num, '(BODY.PEEK[HEADER] BODY.PEEK[TEXT])')
                        if not msg_data or not msg_data[0]:
                            print(f"    ERROR: Failed to fetch email #{num.decode()}")
                            continue
                        # Reconstruct a minimal email from header + text parts
                        header_bytes = b''
                        text_bytes = b''
                        for part in msg_data:
                            if isinstance(part, tuple):
                                if b'HEADER' in part[0]:
                                    header_bytes = part[1]
                                elif b'TEXT' in part[0]:
                                    text_bytes = part[1]
                        email_message = email.message_from_bytes(header_bytes + b'\r\n' + text_bytes)

                        content = self._extract_text_content(email_message)

                        # Decode the To header (may be MIME encoded)
                        raw_to = (email_message['To'] or '').replace('\n', ' ').replace('\r', ' ').strip()
                        decoded_to = self.decode_mime_header(raw_to)
                        recipient = email.utils.parseaddr(decoded_to)[1]

                        raw_subject = (email_message['Subject'] or '').replace('\n', ' ').replace('\r', ' ').strip()
                        subject = self.decode_mime_header(raw_subject)
                        in_reply_to = (email_message.get('In-Reply-To', '') or '').replace('\n', '').replace('\r', '').strip()

                        print(f"\n  Checking sent email:")
                        print(f"    To (raw): {raw_to[:80]}")
                        print(f"    To (decoded): {decoded_to[:80]}")
                        print(f"    To (final): {recipient}")
                        print(f"    Subject (decoded): {subject}")
                        print(f"    In-Reply-To: {in_reply_to}")

                        # Try to match with a pending draft
                        matched_draft = None
                        db_drafts = self._db.execute("SELECT * FROM pending_drafts").fetchall()
                        if not db_drafts:
                            print(f"    ✗ No pending drafts to match against")
                        else:
                            for i, draft in enumerate(db_drafts):
                                print(f"    Checking draft {i+1}/{len(db_drafts)}")
                                print(f"      Draft To: {draft['recipient']}")
                                decoded_draft_subject = self.decode_mime_header(draft['subject'])
                                print(f"      Draft Subject (decoded): {decoded_draft_subject}")
                                print(f"      Draft Message ID: {draft['original_message_id'] or 'N/A'}")

                                if (draft['recipient'].lower() == recipient.lower() and
                                    decoded_draft_subject.lower() in subject.lower()):
                                    print(f"      ✓ Matched by recipient + subject!")
                                    matched_draft = dict(draft)
                                    break
                                elif in_reply_to and draft['original_message_id'] == in_reply_to:
                                    print(f"      ✓ Matched by In-Reply-To header!")
                                    matched_draft = dict(draft)
                                    break
                                else:
                                    print(f"      ✗ No match")

                        if not matched_draft:
                            print(f"    ✗ No matching draft found")

                            # Check if we've already learned from this sent email
                            if message_id and self._db_is_processed(message_id, 'manually_learned'):
                                print(f"    Already learned from this sent email, skipping")
                                continue

                            print(f"    Attempting to find and save original conversation")

                            # Try to find the original email
                            original_email = None

                            # First, try Message-ID search if we have In-Reply-To header
                            if in_reply_to:
                                print(f"    Has In-Reply-To header, searching by Message-ID")
                                original_email = self.find_email_by_message_id(in_reply_to)
                                # CRITICAL: Re-select the Sent folder because find_email_by_message_id() switches folders
                                self.imap.select(folder, readonly=True)

                            # If Message-ID search fails (or no In-Reply-To), try fallback search by recipient + subject
                            if not original_email:
                                print(f"    Searching by recipient + subject")
                                original_email = self.find_email_by_recipient_subject(recipient, subject)
                                # CRITICAL: Re-select the Sent folder because find_email_by_recipient_subject() switches folders
                                self.imap.select(folder, readonly=True)

                            if original_email:
                                print(f"    Found original email from {original_email['sender']}")

                                # Check if the original email would have been filtered (order, spam, etc)
                                if self.is_blacklisted(original_email['sender'], original_email['subject'], original_email['content']):
                                    print(f"    Original email is filtered (order/spam), NOT learning response style")
                                    print(f"    This prevents automatic replies to similar filtered emails")
                                    # BUT: Still save conversation history for context
                                    # For filtered emails (order confirmations), save under the CUSTOMER's email, not shop email
                                    print(f"    Saving conversation history under customer email: {recipient}")
                                    customer_email_data = {
                                        'sender': recipient,  # Use customer's email as sender for conversation history
                                        'subject': original_email['subject'],
                                        'content': original_email['content']
                                    }
                                    self.update_history(customer_email_data, content)

                                    # Track that we processed this email to avoid reprocessing
                                    self._db_mark_processed(message_id, 'manually_learned')
                                    print(f"    Tracked Message-ID to prevent reprocessing")
                                else:
                                    # Save to conversation history
                                    self.update_history(original_email, content)
                                    print(f"    Learned from response")

                                    # Check if this email contains an appointment confirmation
                                    appointment = self.parse_calendar_marker(content, recipient, subject)
                                    if appointment:
                                        print(f"Found appointment in sent email: {appointment['start'].strftime('%d.%m.%Y %H:%M')}")
                                        self.create_calendar_event(appointment)

                                    # Track that we learned from this sent email
                                    self._db_mark_processed(message_id, 'manually_learned')

                                    learned_count += 1
                            else:
                                print(f"    Could not find original email (tried Message-ID and recipient+subject searches)")
                                # Last resort: Save conversation history with just the recipient
                                # This ensures at least we have context for future replies
                                print(f"    Saving basic conversation history with recipient: {recipient}")
                                basic_email_data = {
                                    'sender': recipient,
                                    'subject': subject,
                                    'content': f"[Email - original not found]\n{subject}"
                                }
                                self.update_history(basic_email_data, content)

                                # Check if this email contains an appointment confirmation
                                appointment = self.parse_calendar_marker(content, recipient, subject)
                                if appointment:
                                    print(f"Found appointment in sent email: {appointment['start'].strftime('%d.%m.%Y %H:%M')}")
                                    self.create_calendar_event(appointment)

                                # Track that we processed this email to avoid reprocessing
                                self._db_mark_processed(message_id, 'manually_learned')

                        if matched_draft:
                            # Found a match! Learn from the sent email
                            print(f"Learning from sent email to {recipient}: {subject}")

                            email_data = {
                                'sender': matched_draft['recipient'],
                                'subject': matched_draft['subject'],
                                'content': matched_draft['original_content']
                            }

                            # Save to conversation history
                            self.update_history(email_data, content)
                            print(f"    Updated conversation history for {matched_draft['recipient']}")

                            # Create calendar event from stored appointment data (marker stripped from IMAP draft)
                            cal_raw = matched_draft.get('calendar_appointment')
                            cal_data = json.loads(cal_raw) if cal_raw else None
                            if cal_data:
                                try:
                                    from datetime import datetime as dt
                                    appointment = {
                                        'title': cal_data['title'],
                                        'start': dt.fromisoformat(cal_data['start']),
                                        'end': dt.fromisoformat(cal_data['end']),
                                        'location': cal_data.get('location', ''),
                                        'description': cal_data.get('description', '')
                                    }
                                    print(f"Creating calendar event from draft tracking: {cal_data['title']}")
                                    self.create_calendar_event(appointment)
                                except Exception as e:
                                    print(f"Error creating calendar event: {e}")

                            # Mark as learned
                            self._db_mark_processed(message_id, 'learned')
                            with self._db_lock:
                                self._db.execute("DELETE FROM pending_drafts WHERE id=?", (matched_draft['id'],))
                                self._db.commit()

                            learned_count += 1

                    except Exception as e:
                        print(f"    !!! ERROR processing sent email #{num.decode()}: {str(e)}")
                        import traceback
                        traceback.print_exc()
                        continue

                # Successfully processed this folder
                if learned_count > 0:
                    print(f"Learned from {learned_count} sent emails in {folder}")
                return learned_count

            except Exception as e:
                # This folder doesn't exist or error accessing it
                print(f"Error processing {folder} folder: {e}")
                import traceback
                traceback.print_exc()
                continue

        if learned_count == 0:
            print("No new sent emails to learn from")

        return learned_count

    def _generate_calendar_title(self, email_content: str, recipient: str = "", subject: str = "") -> str:
        """Generate a meaningful calendar event title using Claude based on email context."""
        try:
            # Extract name from email if possible (e.g., "steineval@bluewin.ch" -> "Steineval")
            client_name = recipient.split('@')[0] if recipient else "Client"

            # Create a simple prompt for Claude to generate a title
            prompt = f"""Based on this email, generate a SHORT and MEANINGFUL calendar event title (max 60 chars).
IMPORTANT: Always include the client name/email ({client_name}) in the title.
The title should describe what the appointment is about and who it's with.

Email subject: {subject}
Email to: {recipient}
Email content: {email_content[:300]}

Generate ONLY the title, nothing else. No quotes, no explanation. Example format: "Linux Setup - Steineval"."""

            response = self.anthropic.messages.create(
                model=self.config['claude_model_name'],
                max_tokens=100,
                messages=[
                    {
                        "role": "user",
                        "content": prompt
                    }
                ]
            )

            # Extract the title from response
            title = response.content[0].text.strip() if response.content else "Termin"

            # Ensure it's not too long
            title = title[:60]

            return title if title else "Termin"

        except Exception as e:
            print(f"Error generating calendar title: {e}")
            return "Termin"  # Fallback

    def parse_calendar_marker(self, email_content: str, recipient: str = "", subject: str = "") -> Optional[Dict]:
        """
        Parse calendar marker from email content.
        Returns None if marker was deleted (user reviewed and removed it).
        Returns dict with appointment details if found in the actual email text.
        Handles German date formats with month/day names.
        """
        # Check for new simplified CALENDAR_MARKER format
        # Format: CALENDAR_MARKER|DD.MM.YYYY HH:MM-HH:MM|Title|Location
        cal_match = re.search(r'CALENDAR_MARKER\|(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2})-(\d{2}:\d{2})\|([^|]*)\|?(.*)', email_content)
        if cal_match:
            try:
                date_str, start_time, end_time, title, location = cal_match.groups()
                day, month, year = date_str.split('.')
                s_hour, s_min = start_time.split(':')
                e_hour, e_min = end_time.split(':')
                start = datetime(int(year), int(month), int(day), int(s_hour), int(s_min))
                end = datetime(int(year), int(month), int(day), int(e_hour), int(e_min))

                if start > datetime.now() and start < datetime.now() + timedelta(days=180):
                    title = title.strip() if title.strip() else self._generate_calendar_title(email_content, recipient, subject)
                    return {
                        'start': start, 'end': end,
                        'title': title,
                        'description': email_content[:200],
                        'location': location.strip()
                    }
            except (ValueError, IndexError) as e:
                print(f"Error parsing CALENDAR_MARKER: {e}")

        # Legacy: Check if the old warning marker is still present
        if "DELETE THIS SECTION BEFORE SENDING" in email_content:
            print("Warning: Calendar marker found in sent email - user forgot to delete it. Skipping calendar creation for safety.")
            return None

        # German month name mappings
        german_months = {
            'januar': 1, 'january': 1, 'jänner': 1,
            'februar': 2, 'february': 2,
            'märz': 3, 'march': 3, 'maerz': 3,
            'april': 4,
            'mai': 5, 'may': 5,
            'juni': 6, 'june': 6,
            'juli': 7, 'july': 7,
            'august': 8,
            'september': 9,
            'oktober': 10, 'october': 10,
            'november': 11,
            'dezember': 12, 'december': 12
        }

        # Now parse the actual email content for appointment confirmation
        # Look for date/time patterns in German format
        date_patterns = [
            # With German month name: "3. Dezember um 14:00" or "3. Dezember 14:00"
            r'(\d{1,2})\.\s+([A-Za-z]+)\s+(?:um\s+)?(\d{1,2}):(\d{2})',
            # Numeric format: 28.11.2024 um 14:00
            r'(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(?:um\s+)?(\d{1,2}):(\d{2})',
            # Numeric format: 28.11.24 14:00
            r'(\d{1,2})\.(\d{1,2})\.(\d{2,4})\s+(\d{1,2}):(\d{2})',
            # Numeric format: 28.11. um 14:00 (current year)
            r'(\d{1,2})\.(\d{1,2})\.\s+(?:um\s+)?(\d{1,2}):(\d{2})',
        ]

        appointment = None
        for pattern in date_patterns:
            match = re.search(pattern, email_content)
            if match:
                groups = match.groups()

                try:
                    if len(groups) == 4 and groups[1].isalpha():
                        # German month name format: day, month_name, hour, minute
                        day = int(groups[0])
                        month_name = groups[1].lower()
                        hour = int(groups[2])
                        minute = int(groups[3])

                        # Find the month number
                        month = None
                        for name, num in german_months.items():
                            if name.startswith(month_name[:3]):  # Match first 3 letters
                                month = num
                                break

                        if not month:
                            print(f"Could not parse German month: {month_name}")
                            continue

                        year = datetime.now().year

                    elif len(groups) == 5:
                        # Numeric format with year
                        day, month, year, hour, minute = groups
                        # Handle 2-digit year
                        if len(year) == 2:
                            year = f"20{year}"
                        day, month, hour, minute = int(day), int(month), int(hour), int(minute)
                        year = int(year)

                    elif len(groups) == 4:
                        # Numeric format without year
                        day, month, hour, minute = groups
                        day, month, hour, minute = int(day), int(month), int(hour), int(minute)
                        year = datetime.now().year
                    else:
                        continue

                    start_time = datetime(year, month, day, hour, minute)

                    # Only create appointments for future dates
                    if start_time < datetime.now():
                        print(f"Skipping past appointment: {start_time}")
                        continue

                    # Only create appointments within next 6 months (safety check)
                    if start_time > datetime.now() + timedelta(days=180):
                        print(f"Skipping far-future appointment (>6 months): {start_time}")
                        continue

                    # Default duration: 1 hour
                    end_time = start_time + timedelta(hours=1)

                    # Generate a meaningful title using Claude
                    title = self._generate_calendar_title(email_content, recipient, subject)

                    appointment = {
                        'start': start_time,
                        'end': end_time,
                        'title': title,
                        'description': email_content[:200],  # First 200 chars as description
                        'location': ''  # Can be enhanced later
                    }
                    print(f"Found appointment: {start_time.strftime('%d.%m.%Y %H:%M')} - {title}")
                    return appointment

                except (ValueError, IndexError) as e:
                    print(f"Error parsing date: {e}")
                    continue

        return None

    def create_calendar_event(self, appointment: Dict) -> bool:
        """Create a CalDAV calendar event."""
        if not self.config.get('enable_calendar', False):
            print("Calendar integration is disabled in config")
            return False

        try:
            import caldav
            from icalendar import Calendar, Event as ICalEvent

            # Connect to CalDAV server
            client = caldav.DAVClient(
                url=self.config['caldav_url'],
                username=self.config['caldav_username'],
                password=self.config['caldav_password'],
                timeout=10
            )

            principal = client.principal()
            calendars = principal.calendars()

            # Find the specified calendar
            target_calendar = None
            calendar_name = self.config.get('caldav_calendar', 'Persönlich')

            for cal in calendars:
                if calendar_name in cal.get_display_name():
                    target_calendar = cal
                    break

            if not target_calendar:
                print(f"Calendar '{calendar_name}' not found. Using first available calendar.")
                target_calendar = calendars[0] if calendars else None

            if not target_calendar:
                print("No calendars found!")
                return False

            # Create iCalendar event
            cal = Calendar()
            event = ICalEvent()
            event.add('summary', appointment['title'])
            event.add('dtstart', appointment['start'])
            event.add('dtend', appointment['end'])
            event.add('description', appointment['description'])
            if appointment.get('location'):
                event.add('location', appointment['location'])

            cal.add_component(event)

            # Save event to calendar
            target_calendar.save_event(cal.to_ical())

            print(f"✓ Calendar event created: {appointment['title']} on {appointment['start'].strftime('%d.%m.%Y %H:%M')}")

            # Log the creation
            self._log_calendar_creation(appointment)

            return True

        except Exception as e:
            print(f"Error creating calendar event: {str(e)}")
            import traceback
            print(f"Full error: {traceback.format_exc()}")
            return False

    def _log_calendar_creation(self, appointment: Dict):
        """Log calendar event creation."""
        try:
            with open('calendar_log.json', 'r') as f:
                log = json.load(f)
        except FileNotFoundError:
            log = []

        log.append({
            'timestamp': datetime.now().isoformat(),
            'appointment': {
                'title': appointment['title'],
                'start': appointment['start'].isoformat(),
                'end': appointment['end'].isoformat(),
                'location': appointment.get('location', '')
            }
        })

        with open('calendar_log.json', 'w') as f:
            json.dump(log, f, indent=2)

    def test_caldav_connection(self):
        """Test CalDAV connection on startup (with 10s timeout)."""
        if not self.config.get('enable_calendar', False):
            print("📅 Calendar integration: DISABLED")
            return False

        try:
            import caldav

            print("📅 Testing CalDAV connection...")
            client = caldav.DAVClient(
                url=self.config['caldav_url'],
                username=self.config['caldav_username'],
                password=self.config['caldav_password'],
                timeout=10
            )

            principal = client.principal()
            calendars = principal.calendars()

            calendar_name = self.config.get('caldav_calendar', 'Persönlich')
            target_calendar = None

            for cal in calendars:
                if calendar_name in cal.get_display_name():
                    target_calendar = cal
                    break

            if target_calendar:
                print(f"✅ Calendar connection SUCCESS: '{calendar_name}' calendar found")
                return True
            else:
                print(f"⚠️  Calendar connection OK, but '{calendar_name}' calendar not found")
                print(f"   Available calendars: {[cal.get_display_name() for cal in calendars]}")
                return False

        except Exception as e:
            print(f"❌ Calendar connection FAILED: {str(e)}")
            print(f"   Calendar integration will be disabled")
            return False

    def run(self, interval: int = 300, search_criteria: str = 'UNSEEN'):
        """Run the email assistant with specified check interval."""
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%d-%b-%Y')
        print(f"Email Assistant started. Checking for emails every {interval} seconds.")
        print(f"Only processing emails since: {yesterday}\n")

        # Test CalDAV connection
        self.test_caldav_connection()
        print()

        # Test Matrix connection and start listener thread
        if self._test_matrix_connection():
            matrix_thread = threading.Thread(target=self._matrix_loop, daemon=True)
            matrix_thread.start()
        print()

        while True:
            try:
                print(f"\n{'='*60}")
                print(f"Checking for new emails at {datetime.now()}")
                print(f"{'='*60}\n")


                # First, learn from spam folder
                print("Learning from spam folder...")
                try:
                    self.learn_from_spam_folder()
                except (socket.timeout, imaplib.IMAP4.abort, OSError) as e:
                    print(f"  IMAP error in spam learning: {e}, reconnecting...")
                    self.reconnect_imap()
                print()

                # Learn from sent emails
                print("Learning from sent emails...")
                try:
                    self.learn_from_sent_emails()
                except (socket.timeout, imaplib.IMAP4.abort, OSError) as e:
                    print(f"  IMAP error in sent learning: {e}, reconnecting...")
                    self.reconnect_imap()
                print()

                # Then process new emails
                new_emails = self.get_new_emails(search_criteria)

                if new_emails:
                    print(f"Found {len(new_emails)} new emails to process")
                else:
                    print("No new emails to process")

                for email_data in new_emails:
                    try:
                        # Check if we've already processed this incoming email
                        message_id = email_data.get('message_id', '')
                        if message_id and self._db_is_processed(message_id, 'incoming'):
                            print(f"Email from {email_data['sender']} already processed (ID: {message_id[:20]}...), skipping")
                            continue

                        print(f"Processing email from {email_data['sender']}")

                        # Triage: classify email before generating response
                        triage_enabled = self.config.get('triage_enabled', True)
                        if triage_enabled:
                            triage = self._classify_email(email_data)
                            print(f"  Triage: {triage['category']} (confidence: {triage['confidence']}) - {triage.get('reason', '')}")

                            email_data['triage'] = triage
                            category = triage['category']
                            threshold = self.config.get('triage_confidence_threshold', 0.7)

                            # Low confidence -> needs human review
                            if triage['confidence'] < threshold:
                                print(f"  Low confidence ({triage['confidence']}), escalating to needs_human")
                                category = 'needs_human'

                            # Update contact profile with triage data
                            self._update_contact_from_triage(email_data, triage)

                            if category in ('spam', 'ignore'):
                                print(f"  Triage: {category}, skipping draft generation")
                                if category == 'spam':
                                    self.mark_as_spam(email_data)
                                    self.move_to_junk(email_data['uid'])
                                if message_id:
                                    self._db_mark_processed(message_id, 'incoming')
                                continue

                            if category == 'order_notification':
                                has_conv = self._db_has_conversation(email_data['sender'])
                                if not has_conv:
                                    print(f"  Triage: order_notification, no conversation, skipping")
                                    if message_id:
                                        self._db_mark_processed(message_id, 'incoming')
                                    continue

                        # All real emails -> pending for human decision via Matrix
                        pid = self._save_pending_decision(email_data, triage if triage_enabled else {'category': 'unclassified', 'confidence': 0, 'reason': ''})
                        self._matrix_notify_pending(email_data, triage if triage_enabled else {'category': 'unclassified', 'confidence': 0, 'reason': ''}, pid)
                        if message_id:
                            self._db_mark_processed(message_id, 'incoming')
                        print(f"  Pending [{pid}] for {email_data['sender']}")
                    except Exception as e:
                        print(f"Error processing email from {email_data['sender']}: {str(e)}")
                        import traceback
                        print(f"Full error: {traceback.format_exc()}")
                        continue

                time.sleep(interval)
            except (imaplib.IMAP4.abort, imaplib.IMAP4.error, ConnectionError, OSError) as e:
                print(f"IMAP connection error: {str(e)}")
                # Try to reconnect
                if self.reconnect_imap():
                    print("Continuing after reconnection...")
                    time.sleep(5)  # Short delay before continuing
                else:
                    print("Reconnection failed, waiting 60 seconds before retry...")
                    time.sleep(60)
            except Exception as e:
                print(f"Error occurred in main loop: {str(e)}")
                import traceback
                print(f"Full error: {traceback.format_exc()}")
                time.sleep(60)  # Wait a minute before retrying

if __name__ == "__main__":
    assistant = EmailAssistant()

    # Example of how to add new instructions
    # assistant.add_instruction("Wenn auf technische Fragen geantwortet wird, Codebeispiele beifügen.")

    # Example of how to add a final response for learning
    # email_data = {...}  # Your email data
    # final_response = "Ihre tatsächlich gesendete Antwort"
    # assistant.add_example_response(email_data, final_response)

    assistant.run()
