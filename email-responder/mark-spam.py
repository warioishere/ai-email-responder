#!/usr/bin/env python3
"""
Helper script to mark emails as spam and train the spam filter.
"""

from email_assistant_anthropic import EmailAssistant
import sys

def main():
    print("Email Spam Marker - Mark emails as spam to train the filter")
    print("=" * 60)

    # Initialize the assistant
    assistant = EmailAssistant()

    # Get recent emails (including already read ones)
    print("\nFetching recent emails...")
    recent_emails = assistant.get_new_emails('ALL')

    if not recent_emails:
        print("No emails found.")
        return

    # Show last 10 emails
    display_emails = recent_emails[-10:]
    print(f"\nShowing last {len(display_emails)} emails:\n")

    for idx, email_data in enumerate(display_emails, 1):
        print(f"{idx}. From: {email_data['sender']}")
        print(f"   Subject: {email_data['subject']}")
        print(f"   Preview: {email_data['content'][:100]}...")
        print()

    # Ask user which emails are spam
    while True:
        print("\nEnter email numbers to mark as spam (comma-separated, e.g., 1,3,5)")
        print("Or enter 'q' to quit")

        user_input = input("> ").strip()

        if user_input.lower() == 'q':
            print("Exiting...")
            break

        try:
            # Parse email numbers
            spam_indices = [int(x.strip()) for x in user_input.split(',')]

            # Validate indices
            invalid = [i for i in spam_indices if i < 1 or i > len(display_emails)]
            if invalid:
                print(f"Invalid email numbers: {invalid}")
                continue

            # Mark emails as spam
            for idx in spam_indices:
                email_data = display_emails[idx - 1]
                print(f"\nMarking as spam: {email_data['subject']}")
                assistant.mark_as_spam(email_data)

            print(f"\n✓ Successfully marked {len(spam_indices)} email(s) as spam!")
            print("The filter has learned patterns from these emails.")

            # Ask if user wants to mark more
            again = input("\nMark more emails as spam? (y/n): ").strip().lower()
            if again != 'y':
                break

        except ValueError:
            print("Invalid input. Please enter numbers separated by commas.")
            continue

    print("\nDone! Spam patterns have been saved to learned_spam.json")

if __name__ == "__main__":
    main()
