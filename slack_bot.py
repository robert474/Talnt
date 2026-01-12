#!/usr/bin/env python3
"""
Slack Bot for Resume Formatting
Listens for file uploads in specified channels and automatically formats resumes.
"""

import os
import tempfile
import subprocess
from pathlib import Path
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
import requests

# Get the directory where this script lives
SCRIPT_DIR = Path(__file__).parent.resolve()

# Initialize the Slack app with your bot token
app = App(token=os.environ.get("SLACK_BOT_TOKEN"))

# Supported file types
SUPPORTED_TYPES = ['pdf', 'docx']

def download_file(url, token, dest_path):
    """Download a file from Slack"""
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    with open(dest_path, 'wb') as f:
        f.write(response.content)
    return dest_path

def format_resume_file(input_path, output_path):
    """Run the resume formatter on a file"""
    # Use the existing format_resume.py script
    format_script = SCRIPT_DIR / "format_resume.py"

    result = subprocess.run(
        ['python3', str(format_script), str(input_path)],
        capture_output=True,
        text=True,
        cwd=str(SCRIPT_DIR)
    )

    if result.returncode != 0:
        raise Exception(f"Formatting failed: {result.stderr}")

    return output_path

@app.event("file_shared")
def handle_file_shared(event, client, logger):
    """Handle file upload events"""
    file_id = event.get("file_id")
    channel_id = event.get("channel_id")
    user_id = event.get("user_id")

    try:
        # Get file info
        file_info = client.files_info(file=file_id)
        file_data = file_info["file"]

        filename = file_data.get("name", "")
        file_ext = filename.split('.')[-1].lower() if '.' in filename else ""

        # Check if it's a supported resume file
        if file_ext not in SUPPORTED_TYPES:
            logger.info(f"Skipping unsupported file type: {filename}")
            return

        # Notify user we're processing
        client.chat_postMessage(
            channel=channel_id,
            text=f"Processing resume: *{filename}*... This may take a moment.",
            thread_ts=event.get("event_ts")
        )

        # Create temp directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir = Path(temp_dir)

            # Download the file
            input_path = temp_dir / filename
            download_url = file_data.get("url_private_download")

            if not download_url:
                raise Exception("Could not get download URL for file")

            download_file(
                download_url,
                os.environ.get("SLACK_BOT_TOKEN"),
                input_path
            )

            # Copy to input folder and run formatter
            import shutil
            input_folder = SCRIPT_DIR / "input"
            output_folder = SCRIPT_DIR / "output"

            # Clear input folder and copy new file
            for f in input_folder.glob("*"):
                f.unlink()

            shutil.copy(input_path, input_folder / filename)

            # Run the formatter
            result = subprocess.run(
                ['python3', str(SCRIPT_DIR / "format_resume.py")],
                capture_output=True,
                text=True,
                cwd=str(SCRIPT_DIR)
            )

            if result.returncode != 0:
                raise Exception(f"Formatting failed: {result.stderr}\n{result.stdout}")

            # Find the output file - format_resume.py names output based on
            # the candidate's name extracted from the resume, not the input filename
            # The output format is: {Name}_Formatted.docx (note: capital F)

            # Get list of formatted files, sorted by modification time (newest first)
            formatted_files = sorted(
                output_folder.glob("*_Formatted.docx"),
                key=lambda f: f.stat().st_mtime,
                reverse=True
            )

            if not formatted_files:
                raise Exception("Could not find formatted output file")

            # Use the most recently created formatted file
            output_docx = formatted_files[0]

            # Upload the formatted resume back to Slack
            client.files_upload_v2(
                channel=channel_id,
                file=str(output_docx),
                filename=output_docx.name,
                title=f"Formatted: {output_docx.stem}",
                initial_comment=f"Here's your formatted resume, <@{user_id}>!"
            )

            logger.info(f"Successfully processed {filename}")

    except Exception as e:
        logger.error(f"Error processing file: {e}")
        client.chat_postMessage(
            channel=channel_id,
            text=f"Sorry, there was an error processing the resume: {str(e)}",
            thread_ts=event.get("event_ts")
        )

@app.event("message")
def handle_message(event, logger):
    """Handle message events (required to prevent warnings)"""
    pass

@app.command("/format-resume")
def handle_format_command(ack, respond, command):
    """Handle /format-resume slash command"""
    ack()
    respond(
        text="To format a resume, simply upload a PDF or DOCX file to this channel. "
             "I'll automatically process it and return the formatted version!"
    )

@app.event("app_mention")
def handle_mention(event, client):
    """Handle when someone mentions the bot"""
    client.chat_postMessage(
        channel=event["channel"],
        text="Hi! I'm the Resume Formatter bot. Just upload a PDF or DOCX resume "
             "to this channel and I'll automatically format it for you!",
        thread_ts=event.get("ts")
    )

def main():
    """Start the bot"""
    print("=" * 60)
    print("Resume Formatter Slack Bot")
    print("=" * 60)

    # Check for required environment variables
    bot_token = os.environ.get("SLACK_BOT_TOKEN")
    app_token = os.environ.get("SLACK_APP_TOKEN")

    if not bot_token:
        print("\nError: SLACK_BOT_TOKEN environment variable not set")
        print("\nTo set up the bot:")
        print("1. Go to https://api.slack.com/apps")
        print("2. Create a new app or select existing")
        print("3. Go to 'OAuth & Permissions'")
        print("4. Add these Bot Token Scopes:")
        print("   - channels:history")
        print("   - channels:read")
        print("   - chat:write")
        print("   - files:read")
        print("   - files:write")
        print("   - app_mentions:read")
        print("5. Install the app to your workspace")
        print("6. Copy the Bot User OAuth Token")
        print("\nThen run:")
        print("  export SLACK_BOT_TOKEN='xoxb-your-token'")
        print("  export SLACK_APP_TOKEN='xapp-your-token'")
        print("  python3 slack_bot.py")
        return

    if not app_token:
        print("\nError: SLACK_APP_TOKEN environment variable not set")
        print("\nTo get the App Token:")
        print("1. Go to https://api.slack.com/apps")
        print("2. Select your app")
        print("3. Go to 'Basic Information'")
        print("4. Scroll to 'App-Level Tokens'")
        print("5. Create a token with 'connections:write' scope")
        print("\nThen run:")
        print("  export SLACK_APP_TOKEN='xapp-your-token'")
        return

    print("\nBot is starting...")
    print("Upload a PDF or DOCX resume to any channel where the bot is added!")
    print("\nPress Ctrl+C to stop\n")

    # Start the bot using Socket Mode
    handler = SocketModeHandler(app, app_token)
    handler.start()

if __name__ == "__main__":
    main()
