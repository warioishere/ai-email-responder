# Email Assistant Bot

An intelligent email assistant that processes incoming emails and generates responses using AI (OpenAI GPT or Anthropic Claude). The bot reads emails via IMAP, generates appropriate responses, and saves them as drafts. It features a learning system that improves responses based on your instructions and examples.

## Features

- **Email Processing**:
  - Monitors inbox for new unread emails
  - Filters emails using customizable blacklist
  - Supports different IMAP search criteria
  - Marks emails as read/unread
  - Saves generated responses as drafts

- **AI Integration**:
  - Supports both OpenAI GPT and Anthropic Claude
  - Customizable system prompts
  - Temperature and token limit controls
  - Conversation history tracking

- **Learning Capabilities**:
  - Persistent training context
  - Add new instructions while running
  - Learn from your final edited responses
  - Maintains conversation history per sender

## Prerequisites

- Python 3.7+
- Email account with IMAP access
- API key from OpenAI or Anthropic

## Installation

1. Clone the repository:
```bash
git clone [repository-url]
cd email-assistant
```

2. Install required packages:
```bash
# For OpenAI version
pip install openai pyyaml

# For Claude version
pip install anthropic pyyaml
```

3. Create a configuration file `config.yaml`:
```yaml
# For OpenAI version
email: "your.email@example.com"
password: "your-email-password"
imap_server: "imap.gmail.com"
openai_api_key: "your-openai-api-key"
model_name: "gpt-4-turbo-preview"  # or "gpt-3.5-turbo"
max_tokens: 1000
temperature: 0.7
mark_as_read: true
blacklist:
  - "spam@example.com"
  - "newsletter@"
  - "noreply@"
system_prompt: |
  Your initial system prompt here...

# For Claude version
email: "your.email@example.com"
password: "your-email-password"
imap_server: "imap.gmail.com"
anthropic_api_key: "your-anthropic-api-key"
model_name: "claude-3-opus-20240229"
max_tokens: 1000
temperature: 0.7
mark_as_read: true
blacklist:
  - "spam@example.com"
  - "newsletter@"
  - "noreply@"
system_prompt: |
  Your initial system prompt here...
```

## Usage

### Basic Usage

1. Start the assistant:
```bash
# For OpenAI version
python email_assistant_oai.py

# For Claude version
python email_assistant_claude.py
```

2. The bot will:
   - Check for new unread emails every 5 minutes (configurable)
   - Generate responses using the AI model
   - Save responses as drafts
   - Mark processed emails as read (configurable)

### Advanced Usage

#### Adding New Instructions

You can add new instructions while the bot is running:

```python
assistant = EmailAssistant()
assistant.add_instruction("When responding to technical questions, include code examples.")
```

#### Adding Final Responses for Learning

After editing and sending a final response:

```python
assistant.add_example_response(
    email_data,  # Original email data
    "Your final edited and sent response"
)
```

#### Customizing Search Criteria

You can customize how the bot searches for emails:

```python
# Check only unread emails (default)
assistant.run(search_criteria='UNSEEN')

# Check unread emails from a specific sender
assistant.run(search_criteria='UNSEEN FROM "important@example.com"')

# Check unread emails since a specific date
assistant.run(search_criteria='UNSEEN SINCE "01-Jan-2024"')
```

## File Structure

- `email_assistant_oai.py`: OpenAI version of the assistant
- `email_assistant_claude.py`: Claude version of the assistant
- `config.yaml`: Configuration file
- `training_context.json`: Stores learning context (created automatically)
- `conversation_history.json`: Stores conversation history (created automatically)

## Gmail Setup

For Gmail accounts, you'll need to:
1. Enable IMAP in Gmail settings
2. Create an App Password if using 2FA
3. Use the App Password in your config.yaml

## Security Notes

- Store your API keys and email credentials securely
- Never commit config.yaml with real credentials
- Consider using environment variables for sensitive data
- Review generated responses before sending

# Email Assistant Bot

An intelligent email assistant that processes incoming emails and generates responses using AI (OpenAI GPT or Anthropic Claude). The bot reads emails via IMAP, generates appropriate responses, and saves them as drafts. It features a learning system that improves responses based on your instructions and examples.

[Previous sections remain the same until "Usage"]

## Usage

### Basic Usage

[Previous basic usage section remains the same]

### Service Setup on Ubuntu

1. Create necessary directories and set permissions:
```bash
# Create directories
sudo mkdir -p /opt/email-assistant
sudo mkdir -p /var/log/email-assistant

# Set ownership (replace 'your_username' with your actual username)
sudo chown -R your_username:your_username /opt/email-assistant
sudo chown -R your_username:your_username /var/log/email-assistant
```

2. Create virtual environment and install dependencies:
```bash
python3 -m venv /opt/email-assistant/venv
source /opt/email-assistant/venv/bin/activate
pip install anthropic pyyaml  # or openai for GPT version
```

3. Create environment file for sensitive data:
```bash
sudo nano /opt/email-assistant/.env

# Add these variables:
ANTHROPIC_API_KEY=your_api_key
EMAIL_PASSWORD=your_email_password
```

4. Create systemd service file:
```bash
sudo nano /etc/systemd/system/email-assistant.service
```

Add the following content:
```ini
[Unit]
Description=Email Assistant Service
After=network.target

[Service]
Type=simple
User=your_username
Group=your_username
WorkingDirectory=/opt/email-assistant
Environment="PATH=/opt/email-assistant/venv/bin"
ExecStart=/opt/email-assistant/venv/bin/python /opt/email-assistant/email_assistant_claude.py

# Restart configuration
Restart=always
RestartSec=10
StartLimitIntervalSec=60
StartLimitBurst=3

# Environment file for sensitive data
EnvironmentFile=/opt/email-assistant/.env

[Install]
WantedBy=multi-user.target
```

5. Deploy the service:
```bash
# Copy your code and config
cp email_assistant_claude.py /opt/email-assistant/
cp config.yaml /opt/email-assistant/

# Reload systemd
sudo systemctl daemon-reload

# Enable and start the service
sudo systemctl enable email-assistant
sudo systemctl start email-assistant
```

### Service Management

Common service commands:
```bash
# Check service status
sudo systemctl status email-assistant

# View logs
sudo journalctl -u email-assistant -f

# View application logs
tail -f /var/log/email-assistant/email-assistant.log

# Stop the service
sudo systemctl stop email-assistant

# Restart the service
sudo systemctl restart email-assistant

# Check if service is running
sudo systemctl is-active email-assistant
```

### Updating the Service

To update the code:
```bash
# Stop the service
sudo systemctl stop email-assistant

# Update the code
cp new_email_assistant_claude.py /opt/email-assistant/email_assistant_claude.py

# Start the service
sudo systemctl start email-assistant
```

### Service Features

The systemd service configuration includes:
- Automatic restart on failure
- 10-second delay between restarts
- Maximum of 3 restart attempts within 60 seconds
- Rotating log files (10MB max size, 5 backup files)
- Separate logging for application and system logs
- Environment file for sensitive data
- Run as specific user for security

[Rest of the README remains the same]


## Contributing

Feel free to submit issues and pull requests for:
- Bug fixes
- New features
- Documentation improvements
- Code optimization

## License

MIT License - feel free to use and modify as needed.
