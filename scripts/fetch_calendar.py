import os
import requests
import datetime
import yaml

# Folder where Markdown files will be saved
OUTPUT_DIR = "Learninghub-Event Edition/Live Events"

# Outlook API endpoint
OUTLOOK_API = "https://graph.microsoft.com/v1.0/me/calendar/events"

# Get token from environment variable
TOKEN = os.getenv("OUTLOOK_TOKEN")

def fetch_events():
    headers = {"Authorization": f"Bearer {TOKEN}"}
    response = requests.get(OUTLOOK_API, headers=headers)
    response.raise_for_status()
    return response.json().get("value", [])

def write_event_files(events):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    index_lines = [
        "# 📅 Live Events Calendar\n",
        "| Date | Event | Link |\n",
        "|------|--------|------|\n"
    ]

    for event in events:
        title = event["subject"].replace("/", "-")
        start = event["start"]["dateTime"][:10]
        filename = f"{OUTPUT_DIR}/{title}.md"

        with open(filename, "w", encoding="utf-8") as f:
            f.write(f"# {title}\n\n")
            f.write(f"**Date:** {start}\n")
            f.write(f"**Location:** {event.get('location', {}).get('displayName', 'N/A')}\n\n")
            f.write("## Notes\nAdd your notes here...\n\n")
            f.write("[⬅ Back to Home](README.md)\n")

        index_lines.append(f"| {start} | {title} | [{title}]({title}.md) |\n")

    with open(f"{OUTPUT_DIR}/README.md", "w", encoding="utf-8") as f:
        f.writelines(index_lines)

def main():
    events = fetch_events()
    write_event_files(events)
    print(f"✅ Generated {len(events)} event files.")

if __name__ == "__main__":
    main()
