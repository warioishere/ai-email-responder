import imaplib
import email
import os
from email.message import EmailMessage
from openai import OpenAI
from datetime import datetime
import json
import time
from typing import List, Dict
import yaml

class EmailAssistant:
    def __init__(self, config_path: str = 'config.yaml'):
        """Initialize the email assistant with configuration."""
        self.load_config(config_path)
        self.openai = OpenAI(api_key=self.config['openai_api_key'])
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

    def generate_response(self, email_data: Dict) -> str:
        """Generate response using OpenAI API with full training context."""
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
        
        response = self.openai.chat.completions.create(
            model=self.config['openai_model_name'],
            messages=[
                {"role": "system", "content": full_system_prompt},
                {"role": "user", "content": full_context}
            ],
            max_tokens=self.config.get('max_tokens', 1000),
            temperature=self.config.get('temperature', 0.7)
        )
        
        return response.choices[0].message.content

    # ... (other existing methods remain the same) ...

    def run(self, interval: int = 300, search_criteria: str = 'UNSEEN'):
        """Run the email assistant with specified check interval."""
        while True:
            try:
                print(f"Checking for new emails at {datetime.now()}")
                new_emails = self.get_new_emails(search_criteria)
                
                for email_data in new_emails:
                    print(f"Processing email from {email_data['sender']}")
                    response = self.generate_response(email_data)
                    self.save_draft(email_data, response)
                    self.update_history(email_data, response)
                    
                    if self.config.get('mark_as_read', True):
                        self.mark_as_read(email_data['uid'])
                    
                    print(f"Draft saved for email from {email_data['sender']}")
                
                time.sleep(interval)
            except Exception as e:
                print(f"Error occurred: {e}")
                time.sleep(60)

if __name__ == "__main__":
    assistant = EmailAssistant()
    
    # Example of how to add new instructions
    # assistant.add_instruction("When responding to technical questions, include code examples.")
    
    # Example of how to add a final response for learning
    # email_data = {...}  # Your email data
    # final_response = "Your actual sent response"
    # assistant.add_example_response(email_data, final_response)
    
    assistant.run()
