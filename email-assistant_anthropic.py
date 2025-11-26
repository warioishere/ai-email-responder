import imaplib
import email
import os
from email.message import EmailMessage
from anthropic import Anthropic
from datetime import datetime
import json
import time
from typing import List, Dict
import yaml

class EmailAssistant:
    def __init__(self, config_path: str = 'config.yaml'):
        """Initialize the email assistant with configuration."""
        self.load_config(config_path)
        self.anthropic = Anthropic(api_key=self.config['anthropic_api_key'])
        self.connect_imap()
        self.load_history()
        self.load_training_context()

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

    def connect_imap(self):
        """Connect to IMAP server."""
        self.imap = imaplib.IMAP4_SSL(self.config['imap_server'])
        self.imap.login(self.config['email'], self.config['password'])

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

        return False

    def get_new_emails(self, search_criteria: str = 'UNSEEN') -> List[Dict]:
        """Fetch new emails from inbox."""
        self.imap.select('INBOX')
        _, message_numbers = self.imap.search(None, search_criteria)

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
                        content += part.get_payload(decode=True).decode()
            else:
                content = email_message.get_payload(decode=True).decode()

            sender = email.utils.parseaddr(email_message['From'])[1]
            subject = email_message['Subject'] or ''

            # Check blacklist and filters
            if self.is_blacklisted(sender, subject, content):
                continue

            emails.append({
                'uid': num,
                'sender': sender,
                'subject': subject,
                'content': content,
                'message_id': email_message['Message-ID']
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

    def save_draft(self, email_data: Dict, response: str):
        """Save response as draft email."""
        draft = EmailMessage()
        draft['To'] = email_data['sender']
        draft['Subject'] = f"Re: {email_data['subject']}"
        draft['In-Reply-To'] = email_data['message_id']
        draft.set_content(response)

        # Save to drafts folder
        self.imap.append('Drafts', '', imaplib.Time2Internaldate(time.time()),
                        draft.as_bytes())

    def run(self, interval: int = 300, search_criteria: str = 'UNSEEN'):
        """Run the email assistant with specified check interval."""
        print(f"Email Assistant started. Checking for emails every {interval} seconds.")
        while True:
            try:
                print(f"Checking for new emails at {datetime.now()}")
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
