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

class EmailAssistant:
    def __init__(self, config_path: str = 'config.yaml'):
        """Initialize the email assistant with configuration."""
        self.load_config(config_path)
        self.anthropic = Anthropic(api_key=self.config['anthropic_api_key'])
        self.connect_imap()
        self.load_history()
        self.load_training_context()
        self.load_learned_spam()
        self.load_draft_tracking()

    def load_config(self, config_path: str):
        """Load configuration from YAML file."""
        with open(config_path, 'r') as f:
            self.config = yaml.safe_load(f)

    def load_training_context(self):
        """Load training context from JSON file."""
        try:
            with open('training_context.json', 'r') as f:
                self.training_context = json.load(f)
        except FileNotFoundError:
            self.training_context = {
                'system_prompt': self.config.get('system_prompt', ''),
                'additional_instructions': [],
                'example_responses': []
            }
            self.save_training_context()

    def save_training_context(self):
        """Save training context to JSON file."""
        with open('training_context.json', 'w') as f:
            json.dump(self.training_context, f, indent=2)

    def add_instruction(self, instruction: str):
        """Add a new instruction to the training context."""
        self.training_context['additional_instructions'].append({
            'timestamp': datetime.now().isoformat(),
            'instruction': instruction
        })
        self.save_training_context()
        print(f"Added new instruction to training context")

    def add_example_response(self, email_data: Dict, final_response: str):
        """Add a final response as an example to learn from."""
        self.training_context['example_responses'].append({
            'timestamp': datetime.now().isoformat(),
            'sender': email_data['sender'],
            'subject': email_data['subject'],
            'original_content': email_data['content'],
            'response': final_response
        })
        self.save_training_context()
        print(f"Added final response to training examples")

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
                'manually_sent_learned': []  # Message IDs of manually sent emails we've learned from
            }
            self.save_draft_tracking()

        # Ensure the new field exists in older versions
        if 'manually_sent_learned' not in self.draft_tracking:
            self.draft_tracking['manually_sent_learned'] = []
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

                        # Extract email content
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
                        else:
                            try:
                                payload = email_message.get_payload(decode=True)
                                if payload:
                                    try:
                                        content = payload.decode('utf-8')
                                    except UnicodeDecodeError:
                                        for encoding in ['iso-8859-1', 'windows-1252', 'latin-1']:
                                            try:
                                                content = payload.decode(encoding)
                                                break
                                            except UnicodeDecodeError:
                                                continue
                                        else:
                                            content = payload.decode('utf-8', errors='ignore')
                                else:
                                    content = ''
                            except:
                                content = ''

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

    def load_history(self):
        """Load conversation history from JSON file."""
        try:
            with open('conversation_history.json', 'r') as f:
                self.conversation_history = json.load(f)
        except FileNotFoundError:
            self.conversation_history = {}

    def save_history(self):
        """Save conversation history to JSON file."""
        with open('conversation_history.json', 'w') as f:
            json.dump(self.conversation_history, f, indent=2)

    def _get_relevant_history(self, sender: str) -> str:
        """Get conversation history for a specific sender."""
        if sender not in self.conversation_history:
            return "No previous conversations with this sender."

        history = self.conversation_history[sender]
        history_text = ""
        for conv in history[-3:]:  # Last 3 conversations
            history_text += f"\nDate: {conv['date']}\n"
            history_text += f"Email: {conv['email_content']}\n"
            history_text += f"Response: {conv['response']}\n"

        return history_text

    def update_history(self, email_data: Dict, response: str):
        """Update conversation history with new email and response."""
        sender = email_data['sender']
        if sender not in self.conversation_history:
            self.conversation_history[sender] = []

        self.conversation_history[sender].append({
            'date': datetime.now().isoformat(),
            'email_content': email_data['content'],
            'subject': email_data['subject'],
            'response': response
        })

        self.save_history()

    def is_blacklisted(self, sender: str, subject: str, content: str) -> bool:
        """Check if email should be filtered out based on blacklist or keywords."""
        blacklist = self.config.get('blacklist', [])

        # Check sender blacklist
        for pattern in blacklist:
            if pattern.lower() in sender.lower():
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

            # Extract email content
            if email_message.is_multipart():
                content = ''
                for part in email_message.walk():
                    if part.get_content_type() == "text/plain":
                        try:
                            payload = part.get_payload(decode=True)
                            if payload:
                                # Try UTF-8 first, then fall back to other encodings
                                try:
                                    content += payload.decode('utf-8')
                                except UnicodeDecodeError:
                                    # Try common encodings for German emails
                                    for encoding in ['iso-8859-1', 'windows-1252', 'latin-1']:
                                        try:
                                            content += payload.decode(encoding)
                                            break
                                        except UnicodeDecodeError:
                                            continue
                                    else:
                                        # If all fail, decode with errors='ignore'
                                        content += payload.decode('utf-8', errors='ignore')
                        except:
                            pass
            else:
                try:
                    payload = email_message.get_payload(decode=True)
                    if payload:
                        try:
                            content = payload.decode('utf-8')
                        except UnicodeDecodeError:
                            for encoding in ['iso-8859-1', 'windows-1252', 'latin-1']:
                                try:
                                    content = payload.decode(encoding)
                                    break
                                except UnicodeDecodeError:
                                    continue
                            else:
                                content = payload.decode('utf-8', errors='ignore')
                    else:
                        content = ''
                except:
                    content = ''

            sender = email.utils.parseaddr(email_message['From'])[1]
            raw_subject = (email_message['Subject'] or '').replace('\n', ' ').replace('\r', ' ').strip()
            subject = self.decode_mime_header(raw_subject)
            message_id = (email_message['Message-ID'] or '').replace('\n', '').replace('\r', '').strip()

            # Check blacklist and filters
            if self.is_blacklisted(sender, subject, content):
                continue

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

    def generate_response(self, email_data: Dict) -> str:
        """Generate response using Claude API with full training context."""
        try:
            # Combine all instructions
            full_system_prompt = self.training_context['system_prompt'] + "\n\n"
            if self.training_context['additional_instructions']:
                full_system_prompt += "Additional Instructions:\n"
                for inst in self.training_context['additional_instructions']:
                    full_system_prompt += f"- {inst['instruction']}\n"

            # Create context from examples
            example_context = ""
            if self.training_context['example_responses']:
                recent_examples = sorted(
                    self.training_context['example_responses'],
                    key=lambda x: x['timestamp'],
                    reverse=True
                )[:5]  # Get 5 most recent examples

                example_context = "Recent example responses:\n\n"
                for ex in recent_examples:
                    example_context += f"Subject: {ex['subject']}\n"
                    example_context += f"Original: {ex['original_content']}\n"
                    example_context += f"Response: {ex['response']}\n\n"

            # Current email context
            email_context = f"""
            From: {email_data['sender']}
            Subject: {email_data['subject']}
            Content: {email_data['content']}
            """

            # Get conversation history for this sender
            history_context = self._get_relevant_history(email_data['sender'])

            # Combine all context
            full_context = f"{example_context}\n\nPrevious conversations with this sender:\n{history_context}\n\nNew email to respond to:\n{email_context}"

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

                # Extract email content
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
                else:
                    try:
                        payload = email_message.get_payload(decode=True)
                        if payload:
                            try:
                                content = payload.decode('utf-8')
                            except UnicodeDecodeError:
                                for encoding in ['iso-8859-1', 'windows-1252', 'latin-1']:
                                    try:
                                        content = payload.decode(encoding)
                                        break
                                    except UnicodeDecodeError:
                                        continue
                                else:
                                    content = payload.decode('utf-8', errors='ignore')
                        else:
                            content = ''
                    except:
                        content = ''

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
        """Find original email by searching for order number in subject across multiple folders.
        Used as fallback when Message-ID search fails."""
        try:
            # Extract order number from subject (e.g., "Nr. 19530")
            import re
            order_match = re.search(r'Nr\.\s+(\d+)', subject)
            if not order_match:
                return None

            order_num = order_match.group(1)
            print(f"    Searching for order {order_num} in multiple folders...")

            # Search in multiple folders
            folders_to_check = ['INBOX', 'Archive', 'Sent', 'INBOX.Archive']

            for folder in folders_to_check:
                try:
                    self.imap.select(folder)
                    # Search for emails containing the order number in subject
                    _, result = self.imap.search(None, f'SUBJECT "{order_num}"')

                    if not result[0]:
                        continue

                    email_ids = result[0].split()
                    # Check the most recent emails
                    for email_id in email_ids[-20:]:
                        try:
                            _, msg_data = self.imap.fetch(email_id, '(RFC822)')
                            email_body = msg_data[0][1]
                            email_message = email.message_from_bytes(email_body)

                            # Decode the subject
                            raw_subject = (email_message['Subject'] or '').replace('\n', ' ').replace('\r', ' ').strip()
                            decoded_subject = self.decode_mime_header(raw_subject)

                            # Verify it's the right order by checking order number is in subject
                            if order_num in decoded_subject and 'Bestellung' in decoded_subject:
                                print(f"    Found original email in {folder}: {decoded_subject}")
                                # Extract content
                                content = ''
                                if email_message.is_multipart():
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
                                else:
                                    try:
                                        payload = email_message.get_payload(decode=True)
                                        if payload:
                                            try:
                                                content = payload.decode('utf-8')
                                            except UnicodeDecodeError:
                                                for encoding in ['iso-8859-1', 'windows-1252', 'latin-1']:
                                                    try:
                                                        content = payload.decode(encoding)
                                                        break
                                                    except UnicodeDecodeError:
                                                        continue
                                                else:
                                                    content = payload.decode('utf-8', errors='ignore')
                                        else:
                                            content = ''
                                    except:
                                        content = ''

                                sender = email.utils.parseaddr(email_message['From'])[1]
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
                    continue

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
                # Try to select the sent folder
                status, _ = self.imap.select(folder, readonly=True)
                if status != 'OK':
                    continue  # Folder doesn't exist, try next one

                print(f"Checking {folder} folder for sent emails since {week_ago}...")

                # Get emails sent in last 7 days
                _, message_numbers = self.imap.search(None, f'SINCE {week_ago}')

                if not message_numbers[0]:
                    print(f"No sent emails found in {folder} since {week_ago}")
                    continue  # No emails in this folder

                sent_emails = message_numbers[0].split()
                print(f"Found {len(sent_emails)} sent emails in {folder} to check")
                print(f"Pending drafts to match: {len(self.draft_tracking['pending_drafts'])}")

                for num in sent_emails:
                    try:
                        _, msg_data = self.imap.fetch(num, '(RFC822)')
                        email_body = msg_data[0][1]
                        email_message = email.message_from_bytes(email_body)

                        # Extract email content
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
                        else:
                            try:
                                payload = email_message.get_payload(decode=True)
                                if payload:
                                    try:
                                        content = payload.decode('utf-8')
                                    except UnicodeDecodeError:
                                        for encoding in ['iso-8859-1', 'windows-1252', 'latin-1']:
                                            try:
                                                content = payload.decode(encoding)
                                                break
                                            except UnicodeDecodeError:
                                                continue
                                        else:
                                            content = payload.decode('utf-8', errors='ignore')
                                else:
                                    content = ''
                            except:
                                content = ''

                        recipient = email.utils.parseaddr(email_message['To'])[1]
                        raw_subject = (email_message['Subject'] or '').replace('\n', ' ').replace('\r', ' ').strip()
                        # Decode MIME encoded headers
                        subject = self.decode_mime_header(raw_subject)
                        message_id = (email_message['Message-ID'] or '').replace('\n', '').replace('\r', '').strip()
                        in_reply_to = (email_message.get('In-Reply-To', '') or '').replace('\n', '').replace('\r', '').strip()

                        print(f"\n  Checking sent email:")
                        print(f"    To: {recipient}")
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
                            # Check if this is a manually sent reply (has In-Reply-To but no matching draft)
                            if in_reply_to:
                                print(f"    This appears to be a manually sent reply (has In-Reply-To)")

                                # Check if we've already learned from this manually sent email
                                if message_id in self.draft_tracking['manually_sent_learned']:
                                    print(f"    Already learned from this manually sent email, skipping")
                                    continue

                                # Try to find the original email and learn from it
                                original_email = self.find_email_by_message_id(in_reply_to)

                                # If Message-ID search fails, try fallback search by recipient + subject
                                if not original_email:
                                    print(f"    Message-ID search failed, trying fallback search by recipient+subject")
                                    original_email = self.find_email_by_recipient_subject(recipient, subject)

                                if original_email:
                                    print(f"    Found original email from {original_email['sender']}")

                                    # Check if the original email would have been filtered (order, spam, etc)
                                    if self.is_blacklisted(original_email['sender'], original_email['subject'], original_email['content']):
                                        print(f"    Original email is filtered (order/spam), NOT learning response style")
                                        print(f"    This prevents automatic replies to similar filtered emails")
                                        # BUT: Still save conversation history for context
                                        print(f"    Saving conversation history for context")
                                        self.update_history(original_email, content)
                                    else:
                                        # Only learn if the original email is NOT filtered
                                        self.add_example_response(original_email, content)
                                        print(f"    Learned from manual response")
                                        # Also save to conversation history
                                        self.update_history(original_email, content)

                                        # Track that we learned from this manually sent email
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
                                        'content': f"[Order confirmation or filtered email - original not found]\n{subject}"
                                    }
                                    self.update_history(basic_email_data, content)

                        if matched_draft:
                            # Found a match! Learn from the sent email
                            print(f"Learning from sent email to {recipient}: {subject}")

                            email_data = {
                                'sender': matched_draft['recipient'],
                                'subject': matched_draft['subject'],
                                'content': matched_draft['original_content']
                            }

                            # Add the sent version as a training example
                            self.add_example_response(email_data, content)

                            # Check if this email contains an appointment confirmation
                            appointment = self.parse_calendar_marker(content)
                            if appointment:
                                print(f"Found appointment in sent email: {appointment['start'].strftime('%d.%m.%Y %H:%M')}")
                                self.create_calendar_event(appointment)

                            # Mark as learned
                            self.draft_tracking['learned_from'].append(message_id)
                            self.draft_tracking['pending_drafts'].remove(matched_draft)
                            self.save_draft_tracking()

                            learned_count += 1

                    except Exception as e:
                        print(f"Error processing sent email: {str(e)}")
                        continue

                # Successfully processed this folder
                if learned_count > 0:
                    print(f"Learned from {learned_count} sent emails in {folder}")
                return learned_count

            except Exception as e:
                # This folder doesn't exist or error accessing it
                continue

        if learned_count == 0:
            print("No new sent emails to learn from")

        return learned_count

    def parse_calendar_marker(self, email_content: str) -> Optional[Dict]:
        """
        Parse calendar marker from email content.
        Returns None if marker was deleted (user reviewed and removed it).
        Returns dict with appointment details if found in the actual email text.
        """
        # Check if the warning marker is still present - if yes, user forgot to delete it
        if "⚠️⚠️⚠️ DELETE THIS SECTION BEFORE SENDING" in email_content:
            print("Warning: Calendar marker found in sent email - user forgot to delete it. Skipping calendar creation for safety.")
            return None

        # Now parse the actual email content for appointment confirmation
        # Look for date/time patterns in German format
        date_patterns = [
            r'(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(?:um\s+)?(\d{1,2}):(\d{2})',  # 28.11.2024 um 14:00
            r'(\d{1,2})\.(\d{1,2})\.(\d{2,4})\s+(\d{1,2}):(\d{2})',  # 28.11.24 14:00
            r'(\d{1,2})\.(\d{1,2})\.\s+(?:um\s+)?(\d{1,2}):(\d{2})',  # 28.11. um 14:00 (current year)
        ]

        appointment = None
        for pattern in date_patterns:
            match = re.search(pattern, email_content)
            if match:
                groups = match.groups()
                if len(groups) == 5:
                    day, month, year, hour, minute = groups
                    # Handle 2-digit year
                    if len(year) == 2:
                        year = f"20{year}"
                elif len(groups) == 4:  # Pattern without year
                    day, month, hour, minute = groups
                    year = datetime.now().year

                try:
                    start_time = datetime(int(year), int(month), int(day), int(hour), int(minute))

                    # Only create appointments for future dates
                    if start_time < datetime.now():
                        print(f"Skipping past appointment: {start_time}")
                        return None

                    # Only create appointments within next 6 months (safety check)
                    if start_time > datetime.now() + timedelta(days=180):
                        print(f"Skipping far-future appointment (>6 months): {start_time}")
                        return None

                    # Default duration: 1 hour
                    end_time = start_time + timedelta(hours=1)

                    # Extract title from email (first line or subject-related)
                    lines = [line.strip() for line in email_content.split('\n') if line.strip() and not line.startswith(('Mit freundlichen', 'Mario Hofmann', '---'))]
                    title = lines[0][:50] if lines else "Termin"

                    appointment = {
                        'start': start_time,
                        'end': end_time,
                        'title': title,
                        'description': email_content[:200],  # First 200 chars as description
                        'location': ''  # Can be enhanced later
                    }
                    break
                except (ValueError, IndexError) as e:
                    print(f"Error parsing date: {e}")
                    continue

        return appointment

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
                        print(f"Processing email from {email_data['sender']}")
                        response = self.generate_response(email_data)
                        if response and not response.startswith("Error generating"):
                            self.save_draft(email_data, response)
                            self.update_history(email_data, response)

                            if self.config.get('mark_as_read', True):
                                self.mark_as_read(email_data['uid'])

                            print(f"Draft saved for email from {email_data['sender']}")
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
