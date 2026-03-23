import imaplib
import email
from email.header import decode_header
import os
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
        self.connect_imap()
        self.load_history()
        self.load_learned_spam()
        self.load_draft_tracking()
        self._load_article_index()

    def load_config(self, config_path: str):
        """Load configuration from YAML file."""
        with open(config_path, 'r') as f:
            self.config = yaml.safe_load(f)

    def load_learned_spam(self):
        """Load learned spam patterns from JSON file."""
        try:
            with open('learned_spam.json', 'r') as f:
                self.learned_spam = json.load(f)
        except FileNotFoundError:
            self.learned_spam = {
                'keywords': [],
                'senders': [],
                'spam_emails': [],
                'processed_message_ids': []  # Track which spam emails we've learned from
            }
            self.save_learned_spam()

    def save_learned_spam(self):
        """Save learned spam patterns to JSON file."""
        with open('learned_spam.json', 'w') as f:
            json.dump(self.learned_spam, f, indent=2)

    def load_draft_tracking(self):
        """Load draft tracking data from JSON file."""
        try:
            with open('draft_tracking.json', 'r') as f:
                self.draft_tracking = json.load(f)
        except FileNotFoundError:
            self.draft_tracking = {
                'pending_drafts': [],  # Drafts waiting to be sent
                'learned_from': [],  # Message IDs we've already learned from (drafts)
                'manually_sent_learned': [],  # Message IDs of manually sent emails we've learned from
                'processed_incoming_ids': []  # Message IDs of incoming emails already processed
            }
            self.save_draft_tracking()

        # Ensure the new field exists in older versions
        if 'manually_sent_learned' not in self.draft_tracking:
            self.draft_tracking['manually_sent_learned'] = []
            self.save_draft_tracking()

        if 'processed_incoming_ids' not in self.draft_tracking:
            self.draft_tracking['processed_incoming_ids'] = []
            self.save_draft_tracking()

    def save_draft_tracking(self):
        """Save draft tracking data to JSON file."""
        with open('draft_tracking.json', 'w') as f:
            json.dump(self.draft_tracking, f, indent=2)

    def mark_as_spam(self, email_data: Dict):
        """Mark an email as spam and learn keywords from it."""
        # Check if we've already processed this email
        message_id = email_data.get('message_id', '')
        if message_id and message_id in self.learned_spam['processed_message_ids']:
            return  # Already learned from this email

        # Store the full spam email for reference
        self.learned_spam['spam_emails'].append({
            'timestamp': datetime.now().isoformat(),
            'sender': email_data['sender'],
            'subject': email_data['subject'],
            'content': email_data['content'][:200]  # Store first 200 chars
        })

        # Track that we've processed this email
        if message_id:
            self.learned_spam['processed_message_ids'].append(message_id)

        # Add sender to spam list if not already there
        if email_data['sender'] not in self.learned_spam['senders']:
            self.learned_spam['senders'].append(email_data['sender'])
            print(f"Added {email_data['sender']} to spam sender list")

        # Extract unique phrases from subject (3-5 words)
        subject_words = email_data['subject'].split()
        for i in range(len(subject_words)):
            # Extract 3-word phrases
            if i + 2 < len(subject_words):
                phrase = ' '.join(subject_words[i:i+3])
                if phrase not in self.learned_spam['keywords'] and len(phrase) > 10:
                    self.learned_spam['keywords'].append(phrase)
                    print(f"Learned spam keyword: {phrase}")

        self.save_learned_spam()
        print(f"Email from {email_data['sender']} marked as spam and patterns learned")

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
                        if message_id in self.learned_spam['processed_message_ids']:
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
            self.imap.login(self.config['email'], self.config['password'])
            print("✓ IMAP reconnection successful")
            return True
        except Exception as e:
            print(f"✗ IMAP reconnection failed: {e}")
            return False

    def _sanitize_email_filename(self, email_addr: str) -> str:
        """Convert email@example.com to email_example_com.json"""
        return email_addr.replace('@', '_').replace('.', '_') + '.json'

    def _create_contact(self, email_addr: str) -> Dict:
        """Create a new contact profile."""
        return {
            'email': email_addr,
            'name': email_addr.split('@')[0].replace('.', ' ').replace('_', ' ').title(),
            'category_tags': [],
            'topics': [],
            'interaction_count': 0,
            'first_contact': datetime.now().isoformat(),
            'last_contact': datetime.now().isoformat(),
            'conversations': []
        }

    def _save_contact(self, email_addr: str):
        """Save a single contact file to memory/contacts/."""
        contacts_dir = os.path.join('memory', 'contacts')
        os.makedirs(contacts_dir, exist_ok=True)

        filename = self._sanitize_email_filename(email_addr)
        filepath = os.path.join(contacts_dir, filename)

        contact = self._contacts.get(email_addr, self._create_contact(email_addr))
        contact['conversations'] = self.conversation_history.get(email_addr, [])
        contact['last_contact'] = datetime.now().isoformat()
        contact['interaction_count'] = len(contact['conversations'])

        with open(filepath, 'w') as f:
            json.dump(contact, f, indent=2, ensure_ascii=False)

    def _migrate_conversation_history(self):
        """Migrate flat conversation_history.json to per-contact files under memory/contacts/.
        Runs once, then renames the old file to .bak."""
        print("Migrating conversation_history.json to per-contact files...")
        contacts_dir = os.path.join('memory', 'contacts')
        os.makedirs(contacts_dir, exist_ok=True)
        os.makedirs(os.path.join('memory', 'categories'), exist_ok=True)

        try:
            with open('conversation_history.json', 'r') as f:
                old_history = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            print("No conversation_history.json found or empty, skipping migration")
            return

        migrated = 0
        for email_addr, conversations in old_history.items():
            contact = self._create_contact(email_addr)
            contact['conversations'] = conversations
            contact['interaction_count'] = len(conversations)
            if conversations:
                dates = [conv['date'] for conv in conversations if 'date' in conv]
                if dates:
                    contact['first_contact'] = min(dates)
                    contact['last_contact'] = max(dates)

            filename = self._sanitize_email_filename(email_addr)
            filepath = os.path.join(contacts_dir, filename)
            with open(filepath, 'w') as f:
                json.dump(contact, f, indent=2, ensure_ascii=False)
            migrated += 1

        # Rename old file to .bak
        os.rename('conversation_history.json', 'conversation_history.json.bak')
        print(f"Migration complete: {migrated} contacts migrated to memory/contacts/")
        print("Old file renamed to conversation_history.json.bak")

    def load_history(self):
        """Load conversation history from per-contact files in memory/contacts/."""
        contacts_dir = os.path.join('memory', 'contacts')

        # Migrate if old format exists
        if os.path.exists('conversation_history.json') and not os.path.isdir(contacts_dir):
            self._migrate_conversation_history()

        self.conversation_history = {}
        self._contacts = {}

        if os.path.isdir(contacts_dir):
            for filename in os.listdir(contacts_dir):
                if filename.endswith('.json'):
                    filepath = os.path.join(contacts_dir, filename)
                    try:
                        with open(filepath, 'r') as f:
                            contact = json.load(f)
                        email_addr = contact['email']
                        self.conversation_history[email_addr] = contact.get('conversations', [])
                        self._contacts[email_addr] = contact
                    except (json.JSONDecodeError, KeyError) as e:
                        print(f"Error loading contact file {filename}: {e}")

        # Create directories if they don't exist yet (fresh install)
        os.makedirs(contacts_dir, exist_ok=True)
        os.makedirs(os.path.join('memory', 'categories'), exist_ok=True)

    def save_history(self):
        """Save conversation history. Writes only the last updated contact file."""
        if hasattr(self, '_last_updated_sender') and self._last_updated_sender:
            self._save_contact(self._last_updated_sender)
            self._last_updated_sender = None

    def _get_relevant_history(self, sender: str) -> str:
        """Get conversation history for a specific sender.
        Also cleans up conversations older than 14 weeks."""
        if sender not in self.conversation_history:
            return "No previous conversations with this sender."

        # Clean up old conversations (older than 14 weeks = 98 days)
        self._cleanup_old_conversations(sender, weeks=14)

        history = self.conversation_history[sender]
        history_text = ""
        for conv in history[-3:]:  # Last 3 conversations
            history_text += f"\nDate: {conv['date']}\n"
            history_text += f"Email: {conv['email_content']}\n"
            history_text += f"Response: {conv['response']}\n"

        return history_text

    def _get_recent_emails_context(self, days: int = 10) -> str:
        """Build a summary of all conversations from the last N days for global context.
        Helps the assistant understand the broader email landscape."""
        cutoff = datetime.now() - timedelta(days=days)
        recent = []

        for sender, conversations in self.conversation_history.items():
            for conv in conversations:
                try:
                    conv_date = datetime.fromisoformat(conv['date'])
                    if conv_date > cutoff:
                        recent.append({
                            'date': conv['date'],
                            'sender': sender,
                            'subject': conv.get('subject', ''),
                            'snippet': conv.get('email_content', '')[:150],
                            'responded': bool(conv.get('response', ''))
                        })
                except (ValueError, KeyError):
                    continue

        if not recent:
            return ""

        # Sort by date, most recent first, limit to 20
        recent.sort(key=lambda x: x['date'], reverse=True)
        recent = recent[:20]

        context = f"Recent email activity (last {days} days, {len(recent)} emails):\n"
        for item in recent:
            status = "replied" if item['responded'] else "unread"
            context += f"- {item['date'][:10]} | {item['sender']} | {item['subject'][:60]} [{status}]\n"

        return context

    def _cleanup_old_conversations(self, sender: str, weeks: int = 14):
        """Remove conversations older than the specified number of weeks."""
        if sender not in self.conversation_history:
            return

        cutoff_date = datetime.now() - timedelta(weeks=weeks)
        original_count = len(self.conversation_history[sender])

        # Keep only conversations after the cutoff date
        self.conversation_history[sender] = [
            conv for conv in self.conversation_history[sender]
            if datetime.fromisoformat(conv['date']) > cutoff_date
        ]

        removed_count = original_count - len(self.conversation_history[sender])
        if removed_count > 0:
            print(f"Cleaned up {removed_count} old conversations for {sender}")
            self._save_contact(sender)

        # If all conversations were removed, delete the sender entry
        if not self.conversation_history[sender]:
            del self.conversation_history[sender]
            # Remove contact file too
            contacts_dir = os.path.join('memory', 'contacts')
            filepath = os.path.join(contacts_dir, self._sanitize_email_filename(sender))
            if os.path.exists(filepath):
                os.remove(filepath)
            if sender in self._contacts:
                del self._contacts[sender]

    def update_history(self, email_data: Dict, response: str):
        """Update conversation history with new email and response."""
        sender = email_data['sender']
        if sender not in self.conversation_history:
            self.conversation_history[sender] = []
        if sender not in self._contacts:
            self._contacts[sender] = self._create_contact(sender)

        self.conversation_history[sender].append({
            'date': datetime.now().isoformat(),
            'email_content': email_data['content'],
            'subject': email_data['subject'],
            'response': response
        })

        self._last_updated_sender = sender
        self.save_history()

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
        if sender in self.learned_spam['senders']:
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
        for keyword in self.learned_spam['keywords']:
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
            _, msg_data = self.imap.fetch(num, '(RFC822)')
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
            has_conversation = sender in self.conversation_history and len(self.conversation_history[sender]) > 0

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
            has_history = email_data['sender'] in self.conversation_history and \
                          len(self.conversation_history[email_data['sender']]) > 0

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
            return {'category': 'needs_human', 'confidence': 0.0, 'reason': f'Classification error: {str(e)}'}

    def _save_pending_decision(self, email_data: Dict, triage: Dict):
        """Save an email that needs human decision to memory/categories/pending_decisions.json"""
        categories_dir = os.path.join('memory', 'categories')
        os.makedirs(categories_dir, exist_ok=True)
        filepath = os.path.join(categories_dir, 'pending_decisions.json')

        try:
            with open(filepath, 'r') as f:
                decisions = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            decisions = []

        decisions.append({
            'timestamp': datetime.now().isoformat(),
            'sender': email_data['sender'],
            'subject': email_data['subject'],
            'content_preview': email_data['content'][:300] if email_data['content'] else '',
            'message_id': email_data.get('message_id', ''),
            'triage_category': triage['category'],
            'triage_confidence': triage['confidence'],
            'triage_reason': triage.get('reason', ''),
            'resolved': False
        })

        with open(filepath, 'w') as f:
            json.dump(decisions, f, indent=2, ensure_ascii=False)

        print(f"  Saved to pending decisions: {triage['category']} - {triage.get('reason', '')}")

    def _update_contact_from_triage(self, email_data: Dict, triage: Dict):
        """Update contact profile with triage information (category tags, topics)."""
        sender = email_data['sender']
        if sender not in self._contacts:
            self._contacts[sender] = self._create_contact(sender)

        contact = self._contacts[sender]

        # Add category tag if not already present
        category = triage['category']
        if category in ('quick_answer', 'paid_consultation') and category not in contact['category_tags']:
            contact['category_tags'].append(category)

        # Extract topic from reason if meaningful
        reason = triage.get('reason', '')
        if reason and len(reason) > 3:
            # Keep topics list to max 10, avoid duplicates
            if reason not in contact['topics'] and len(contact['topics']) < 10:
                contact['topics'].append(reason)

        self._contacts[sender] = contact
        self._save_contact(sender)

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

    def _matrix_notify_pending(self, email_data: Dict, triage: Dict):
        """Send a Matrix notification for emails that need human review."""
        # Get the number of this pending decision
        try:
            cat_dir = os.path.join('memory', 'categories')
            with open(os.path.join(cat_dir, 'pending_decisions.json'), 'r') as f:
                all_decisions = json.load(f)
            pending_num = len([d for d in all_decisions if not d.get('resolved', False)])
        except:
            pending_num = '?'

        category = triage['category']
        emoji = {"needs_human": "\u2753", "appointment_request": "\U0001f4c5"}.get(category, "\U0001f4e7")
        sender = email_data['sender']
        subject = email_data['subject']
        reason = triage.get('reason', '')

        text = f"""{emoji} [{pending_num}] {category.upper()}
Von: {sender}
Betreff: {subject}
Grund: {reason}

-> draft | call | zeit [wann] | antwort [text] | spam | ignore"""

        html = f"""<b>{emoji} [{pending_num}] {category.upper()}</b><br/>
<b>Von:</b> {sender}<br/>
<b>Betreff:</b> {subject}<br/>
<b>Grund:</b> {reason}<br/>
<br/>
<code>draft</code> | <code>call</code> | <code>zeit [wann]</code> | <code>antwort [text]</code> | <code>spam</code> | <code>ignore</code>"""

        if self._matrix_send_message(text, html):
            print(f"  Matrix notification sent for {sender}")
        else:
            print(f"  Failed to send Matrix notification for {sender}")

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

        room_id = self.config.get('matrix_room_id', '')

        # Long-poll: server holds connection until new events or 30s timeout
        result = self._matrix_sync(timeout=30000)
        if not result:
            return

        # Save sync token
        next_batch = result.get('next_batch')
        if next_batch:
            self._matrix_since = next_batch
            self._save_matrix_token()

        # Process messages from the room
        rooms = result.get('rooms', {}).get('join', {})
        room_data = rooms.get(room_id, {})
        events = room_data.get('timeline', {}).get('events', [])

        # Get the bot's own user ID to ignore its own messages
        bot_user_id = None
        whoami = self._matrix_request('GET', '/account/whoami')
        if whoami:
            bot_user_id = whoami.get('user_id')

        for event in events:
            if event.get('type') != 'm.room.message':
                continue
            if event.get('sender') == bot_user_id:
                continue

            body = event.get('content', {}).get('body', '').strip().lower()
            if not body:
                continue

            # Help command - no pending decision needed
            if body == '!help':
                pending_count = 0
                try:
                    cat_dir = os.path.join('memory', 'categories')
                    with open(os.path.join(cat_dir, 'pending_decisions.json'), 'r') as f:
                        pending_count = len([d for d in json.load(f) if not d.get('resolved', False)])
                except:
                    pass

                help_html = f"""<h4>📧 Email-Assistent</h4>
<b>Model:</b> <code>{self.config['claude_model_name']}</code><br/>
<b>Offene Entscheidungen:</b> {pending_count}<br/>
<br/>
<table>
<tr><td><code>draft</code></td><td>Draft generieren lassen</td></tr>
<tr><td><code>call</code></td><td>Beratungsangebot (135 CHF/h)</td></tr>
<tr><td><code>zeit [wann]</code></td><td>Terminvorschlag senden</td></tr>
<tr><td><code>antwort [text]</code></td><td>Draft mit deinen Anweisungen</td></tr>
<tr><td><code>spam</code></td><td>Als Spam markieren + lernen</td></tr>
<tr><td><code>ignore</code></td><td>Nichts tun</td></tr>
<tr><td><code>!status</code></td><td>Offene Emails anzeigen</td></tr>
<tr><td><code>!help</code></td><td>Diese Hilfe</td></tr>
</table>
<br/><i>Nummer voranstellen fuer gezieltes Targeting: <code>1 spam</code>, <code>2 draft</code></i>"""
                self._matrix_send_html(help_html)
                continue

            # Status command
            if body == '!status':
                try:
                    cat_dir = os.path.join('memory', 'categories')
                    with open(os.path.join(cat_dir, 'pending_decisions.json'), 'r') as f:
                        decisions_list = json.load(f)
                    pending_list = [d for d in decisions_list if not d.get('resolved', False)]
                    if not pending_list:
                        self._matrix_send_html("<i>Keine offenen Entscheidungen.</i>")
                    else:
                        status_html = f"<b>{len(pending_list)} offene Email(s):</b><br/><br/>"
                        for i, d in enumerate(pending_list, 1):
                            cat = d.get('triage_category', '?')
                            reason = d.get('triage_reason', '')
                            status_html += f"<b>{i}.</b> {d['sender']}<br/>"
                            status_html += f"&nbsp;&nbsp;&nbsp;📋 {d['subject']}<br/>"
                            status_html += f"&nbsp;&nbsp;&nbsp;<code>[{cat}]</code> <i>{reason}</i><br/><br/>"
                        self._matrix_send_html(status_html)
                except:
                    self._matrix_send_html("<i>Keine offenen Entscheidungen.</i>")
                continue

            # Load pending decisions
            categories_dir = os.path.join('memory', 'categories')
            filepath = os.path.join(categories_dir, 'pending_decisions.json')
            try:
                with open(filepath, 'r') as f:
                    decisions = json.load(f)
            except (FileNotFoundError, json.JSONDecodeError):
                continue

            # Find unresolved decisions
            pending = [(i, d) for i, d in enumerate(decisions) if not d.get('resolved', False)]
            if not pending:
                self._matrix_send_html("<i>Keine offenen Entscheidungen.</i>")
                continue

            # Parse optional number prefix: "2 draft", "3 spam", or just "draft" (= last)
            original_body = event.get('content', {}).get('body', '').strip()
            target_idx = None
            command = body
            parts = body.split(' ', 1)
            if len(parts) >= 2 and parts[0].isdigit():
                target_idx = int(parts[0])
                command = parts[1].strip()
                # For 'antwort' and 'zeit', preserve original case for the text part
                original_parts = original_body.split(' ', 2)
                if len(original_parts) >= 3:
                    original_body = original_parts[2]  # text after "N command"
                elif len(original_parts) >= 2:
                    original_body = original_parts[1]

            # Select target decision
            if target_idx is not None and 1 <= target_idx <= len(pending):
                _, target = pending[target_idx - 1]
            else:
                _, target = pending[-1]  # default: most recent

            if command == 'draft':
                self._handle_matrix_draft(target, decisions, filepath)
            elif command == 'call':
                self._handle_matrix_call(target, decisions, filepath)
            elif command.startswith('zeit '):
                time_str = original_body[5:] if not parts[0].isdigit() else original_body
                self._handle_matrix_appointment(target, time_str, decisions, filepath)
            elif command == 'ignore':
                target['resolved'] = True
                with open(filepath, 'w') as f:
                    json.dump(decisions, f, indent=2, ensure_ascii=False)
                self._matrix_send_html(f"🚫 <b>Ignoriert:</b> {target['sender']}<br/><i>{target['subject']}</i>")
            elif command == 'spam':
                spam_data = {
                    'sender': target['sender'],
                    'subject': target['subject'],
                    'content': target.get('content_preview', ''),
                    'message_id': target.get('message_id', '')
                }
                self.mark_as_spam(spam_data)
                target['resolved'] = True
                with open(filepath, 'w') as f:
                    json.dump(decisions, f, indent=2, ensure_ascii=False)
                self._matrix_send_html(f"🗑️ <b>Spam gelernt:</b> {target['sender']}<br/><i>Absender und Keywords werden ab jetzt gefiltert.</i>")
            elif command.startswith('antwort '):
                instructions = original_body[8:] if not parts[0].isdigit() else original_body
                self._handle_matrix_custom(target, instructions, decisions, filepath)

    def _handle_matrix_draft(self, decision: Dict, decisions: list, filepath: str):
        """Generate and save a draft for a pending decision."""
        email_data = {
            'sender': decision['sender'],
            'subject': decision['subject'],
            'content': decision.get('content_preview', ''),
            'message_id': decision.get('message_id', ''),
            'triage': {'category': 'quick_answer'}
        }
        response = self.generate_response(email_data)
        if response and not response.startswith("Error"):
            self.save_draft(email_data, response)
            self.update_history(email_data, response)
            decision['resolved'] = True
            with open(filepath, 'w') as f:
                json.dump(decisions, f, indent=2, ensure_ascii=False)
            self._matrix_send_html(f"✅ <b>Draft erstellt</b><br/>An: {decision['sender']}<br/>📋 <i>{decision['subject']}</i>")
        else:
            self._matrix_send_html(f"❌ <b>Fehler:</b> {response}")

    def _handle_matrix_call(self, decision: Dict, decisions: list, filepath: str):
        """Generate a paid consultation response for a pending decision."""
        email_data = {
            'sender': decision['sender'],
            'subject': decision['subject'],
            'content': decision.get('content_preview', ''),
            'message_id': decision.get('message_id', ''),
            'triage': {'category': 'paid_consultation'}
        }
        response = self.generate_response(email_data)
        if response and not response.startswith("Error"):
            self.save_draft(email_data, response)
            self.update_history(email_data, response)
            decision['resolved'] = True
            with open(filepath, 'w') as f:
                json.dump(decisions, f, indent=2, ensure_ascii=False)
            self._matrix_send_html(f"✅ <b>Beratungsangebot-Draft erstellt</b><br/>An: {decision['sender']}<br/>💰 <i>135 CHF/h</i>")
        else:
            self._matrix_send_html(f"❌ <b>Fehler:</b> {response}")

    def _handle_matrix_appointment(self, decision: Dict, time_str: str, decisions: list, filepath: str):
        """Generate an appointment proposal response."""
        email_data = {
            'sender': decision['sender'],
            'subject': decision['subject'],
            'content': decision.get('content_preview', ''),
            'message_id': decision.get('message_id', ''),
            'triage': {'category': 'quick_answer'}
        }
        # Add appointment time to the context so Claude includes it in the response
        email_data['content'] += f"\n\n[SYSTEM: Der Kunde hat nach einem Termin gefragt. Mario hat Zeit am: {time_str}. Schlage diesen Termin vor.]"

        response = self.generate_response(email_data)
        if response and not response.startswith("Error"):
            self.save_draft(email_data, response)
            self.update_history(email_data, response)
            decision['resolved'] = True
            with open(filepath, 'w') as f:
                json.dump(decisions, f, indent=2, ensure_ascii=False)
            self._matrix_send_html(f"✅ <b>Terminvorschlag-Draft erstellt</b><br/>An: {decision['sender']}<br/>📅 <i>{time_str}</i>")
        else:
            self._matrix_send_html(f"❌ <b>Fehler:</b> {response}")

    def _handle_matrix_custom(self, decision: Dict, instructions: str, decisions: list, filepath: str):
        """Generate a draft based on custom instructions from Matrix."""
        email_data = {
            'sender': decision['sender'],
            'subject': decision['subject'],
            'content': decision.get('content_preview', ''),
            'message_id': decision.get('message_id', ''),
            'triage': {'category': 'quick_answer'}
        }
        email_data['content'] += f"\n\n[SYSTEM: Mario moechte folgendes antworten. Formuliere seine Anweisungen als professionelle, freundliche Email in der Sprache des Kunden: {instructions}]"

        response = self.generate_response(email_data)
        if response and not response.startswith("Error"):
            self.save_draft(email_data, response)
            self.update_history(email_data, response)
            decision['resolved'] = True
            with open(filepath, 'w') as f:
                json.dump(decisions, f, indent=2, ensure_ascii=False)
            self._matrix_send_html(f"✅ <b>Draft erstellt</b> (nach deinen Anweisungen)<br/>An: {decision['sender']}<br/>📋 <i>{decision['subject']}</i>")
        else:
            self._matrix_send_html(f"❌ <b>Fehler:</b> {response}")

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
        """Find relevant articles based on email subject and content using keyword matching."""
        if not self._article_index:
            return ""

        subject = email_data.get('subject', '').lower()
        content = email_data.get('content', '')[:500].lower()
        search_text = f"{subject} {content}"

        # Score each article by keyword overlap
        scored = []
        for article in self._article_index:
            title_lower = article['title'].lower()
            title_words = set(re.findall(r'\w{4,}', title_lower))
            search_words = set(re.findall(r'\w{4,}', search_text))

            # Count matching words
            overlap = title_words & search_words
            # Boost for category match
            cat_text = ' '.join(article.get('categories', [])).lower()
            cat_words = set(re.findall(r'\w{4,}', cat_text))
            cat_overlap = cat_words & search_words

            score = len(overlap) * 2 + len(cat_overlap)
            if score > 0:
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

        # Remove markdown formatting from response
        clean_response = self.remove_markdown(response)

        draft = EmailMessage()
        draft['To'] = email_data['sender']
        draft['Subject'] = f"Re: {subject}"
        if message_id:
            draft['In-Reply-To'] = message_id
        draft.set_content(clean_response)

        # Save to drafts folder
        self.imap.append('Drafts', '', imaplib.Time2Internaldate(time.time()),
                        draft.as_bytes())

        # Track this draft for learning
        self.draft_tracking['pending_drafts'].append({
            'timestamp': datetime.now().isoformat(),
            'recipient': email_data['sender'],
            'subject': email_data['subject'],
            'original_content': email_data['content'],
            'draft_response': response,
            'original_message_id': email_data['message_id']
        })
        self.save_draft_tracking()

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

        # Clean up old pending drafts (older than 30 days - probably never sent)
        cutoff_date = datetime.now() - timedelta(days=30)
        old_count = len(self.draft_tracking['pending_drafts'])
        self.draft_tracking['pending_drafts'] = [
            draft for draft in self.draft_tracking['pending_drafts']
            if datetime.fromisoformat(draft['timestamp']) > cutoff_date
        ]
        new_count = len(self.draft_tracking['pending_drafts'])
        if old_count > new_count:
            print(f"Cleaned up {old_count - new_count} old pending drafts (>30 days)")
            self.save_draft_tracking()

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
                print(f"Found {len(sent_emails)} sent emails in {folder} to check")
                print(f"Pending drafts to match: {len(self.draft_tracking['pending_drafts'])}")

                for idx, num in enumerate(sent_emails):
                    try:
                        print(f"  [{idx+1}/{len(sent_emails)}] Fetching email ID {num.decode()}...")
                        _, msg_data = self.imap.fetch(num, '(RFC822)')
                        if not msg_data or not msg_data[0] or len(msg_data[0]) < 2:
                            print(f"    ERROR: Failed to fetch email #{num.decode()}")
                            continue
                        email_body = msg_data[0][1]
                        email_message = email.message_from_bytes(email_body)

                        # Debug: Show subject immediately to find 19530
                        try:
                            raw_subject_debug = (email_message['Subject'] or '').replace('\n', ' ').replace('\r', ' ').strip()
                            subject_debug = self.decode_mime_header(raw_subject_debug)
                            print(f"    Decoded subject (FULL): {subject_debug}")
                            if '19530' in subject_debug or '19530' in raw_subject_debug:
                                print(f"    >>> FOUND 19530 IN EMAIL #{num.decode()}")
                                raw_to_debug = (email_message['To'] or '').replace('\n', ' ').replace('\r', ' ').strip()
                                decoded_to_debug = self.decode_mime_header(raw_to_debug)
                                print(f"    >>> To (decoded): {decoded_to_debug}")
                        except Exception as e:
                            print(f"    ERROR in debug code: {e}")
                            import traceback
                            traceback.print_exc()

                        content = self._extract_text_content(email_message)

                        # Decode the To header (may be MIME encoded)
                        raw_to = (email_message['To'] or '').replace('\n', ' ').replace('\r', ' ').strip()
                        decoded_to = self.decode_mime_header(raw_to)
                        recipient = email.utils.parseaddr(decoded_to)[1]

                        raw_subject = (email_message['Subject'] or '').replace('\n', ' ').replace('\r', ' ').strip()
                        # Decode MIME encoded headers
                        subject = self.decode_mime_header(raw_subject)
                        message_id = (email_message['Message-ID'] or '').replace('\n', '').replace('\r', '').strip()
                        in_reply_to = (email_message.get('In-Reply-To', '') or '').replace('\n', '').replace('\r', '').strip()

                        print(f"\n  Checking sent email:")
                        print(f"    To (raw): {raw_to[:80]}")  # First 80 chars of raw To
                        print(f"    To (decoded): {decoded_to[:80]}")
                        print(f"    To (final): {recipient}")
                        print(f"    Subject (decoded): {subject}")
                        print(f"    In-Reply-To: {in_reply_to}")

                        # Check if we've already learned from this email
                        if message_id in self.draft_tracking['learned_from']:
                            print(f"    Already learned from this email, skipping")
                            continue

                        # Try to match with a pending draft
                        matched_draft = None
                        if not self.draft_tracking['pending_drafts']:
                            print(f"    ✗ No pending drafts to match against")
                        else:
                            for i, draft in enumerate(self.draft_tracking['pending_drafts']):
                                print(f"    Checking draft {i+1}/{len(self.draft_tracking['pending_drafts'])}")
                                print(f"      Draft To: {draft['recipient']}")
                                # Decode draft subject too (it might be encoded)
                                decoded_draft_subject = self.decode_mime_header(draft['subject'])
                                print(f"      Draft Subject (decoded): {decoded_draft_subject}")
                                print(f"      Draft Message ID: {draft.get('original_message_id', 'N/A')}")

                                # Match by recipient and subject similarity
                                if (draft['recipient'].lower() == recipient.lower() and
                                    decoded_draft_subject.lower() in subject.lower()):
                                    print(f"      ✓ Matched by recipient + subject!")
                                    matched_draft = draft
                                    break
                                # Or match by In-Reply-To header
                                elif in_reply_to and draft.get('original_message_id') == in_reply_to:
                                    print(f"      ✓ Matched by In-Reply-To header!")
                                    matched_draft = draft
                                    break
                                else:
                                    print(f"      ✗ No match")

                        if not matched_draft:
                            print(f"    ✗ No matching draft found")

                            # Check if we've already learned from this sent email
                            if message_id in self.draft_tracking['manually_sent_learned']:
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
                                    self.draft_tracking['manually_sent_learned'].append(message_id)
                                    self.save_draft_tracking()
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
                                    self.draft_tracking['manually_sent_learned'].append(message_id)
                                    self.save_draft_tracking()

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
                                self.draft_tracking['manually_sent_learned'].append(message_id)
                                self.save_draft_tracking()

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

                            # Check if this email contains an appointment confirmation
                            appointment = self.parse_calendar_marker(content, matched_draft['recipient'], matched_draft['subject'])
                            if appointment:
                                print(f"Found appointment in sent email: {appointment['start'].strftime('%d.%m.%Y %H:%M')}")
                                self.create_calendar_event(appointment)

                            # Mark as learned
                            self.draft_tracking['learned_from'].append(message_id)
                            self.draft_tracking['pending_drafts'].remove(matched_draft)
                            self.save_draft_tracking()

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
                password=self.config['caldav_password']
            )

            principal = client.principal()
            calendars = principal.calendars()

            # Find the specified calendar
            target_calendar = None
            calendar_name = self.config.get('caldav_calendar', 'Persönlich')

            for cal in calendars:
                if calendar_name in cal.name:
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
        """Test CalDAV connection on startup."""
        if not self.config.get('enable_calendar', False):
            print("📅 Calendar integration: DISABLED")
            return False

        try:
            import caldav

            print("📅 Testing CalDAV connection...")
            client = caldav.DAVClient(
                url=self.config['caldav_url'],
                username=self.config['caldav_username'],
                password=self.config['caldav_password']
            )

            principal = client.principal()
            calendars = principal.calendars()

            calendar_name = self.config.get('caldav_calendar', 'Persönlich')
            target_calendar = None

            for cal in calendars:
                if calendar_name in cal.name:
                    target_calendar = cal
                    break

            if target_calendar:
                print(f"✅ Calendar connection SUCCESS: '{calendar_name}' calendar found")
                print(f"   URL: {self.config['caldav_url']}")
                print(f"   User: {self.config['caldav_username']}")
                return True
            else:
                print(f"⚠️  Calendar connection OK, but '{calendar_name}' calendar not found")
                print(f"   Available calendars: {[cal.name for cal in calendars]}")
                return False

        except Exception as e:
            print(f"❌ Calendar connection FAILED: {str(e)}")
            print(f"   URL: {self.config['caldav_url']}")
            print(f"   User: {self.config['caldav_username']}")
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
                self.learn_from_spam_folder()
                print()

                # Learn from sent emails
                print("Learning from sent emails...")
                self.learn_from_sent_emails()
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
                        if message_id and message_id in self.draft_tracking['processed_incoming_ids']:
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
                                    self.draft_tracking['processed_incoming_ids'].append(message_id)
                                    self.save_draft_tracking()
                                continue

                            if category in ('needs_human', 'appointment_request'):
                                print(f"  Triage: {category}, saving for human review")
                                self._save_pending_decision(email_data, triage)
                                self._matrix_notify_pending(email_data, triage)
                                if message_id:
                                    self.draft_tracking['processed_incoming_ids'].append(message_id)
                                    self.save_draft_tracking()
                                continue

                            if category == 'order_notification':
                                # Order/sale notifications: don't auto-reply unless there's an existing conversation
                                has_conv = email_data['sender'] in self.conversation_history and \
                                           len(self.conversation_history[email_data['sender']]) > 0
                                if has_conv:
                                    print(f"  Triage: order_notification but existing conversation found, generating draft")
                                    # Fall through to generate response
                                else:
                                    print(f"  Triage: order_notification, no existing conversation, skipping")
                                    if message_id:
                                        self.draft_tracking['processed_incoming_ids'].append(message_id)
                                        self.save_draft_tracking()
                                    continue

                        # quick_answer or paid_consultation (or triage disabled) -> generate draft
                        response = self.generate_response(email_data)
                        if response and not response.startswith("Error generating"):
                            self.save_draft(email_data, response)
                            self.update_history(email_data, response)

                            # Track that we've processed this incoming email
                            if message_id:
                                self.draft_tracking['processed_incoming_ids'].append(message_id)
                                self.save_draft_tracking()

                            # Never mark as read - prevents emails from going unnoticed

                            category_info = f" [{email_data.get('triage', {}).get('category', 'unclassified')}]" if triage_enabled else ""
                            print(f"Draft saved for email from {email_data['sender']}{category_info}")
                        else:
                            print(f"Skipping draft save due to error for {email_data['sender']}")
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
