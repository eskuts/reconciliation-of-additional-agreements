# Public Contract Event Monitor (Telegram Bot)

I built this tool to automate the tracking of updates in public procurement contracts. Instead of manually checking for updates, the system monitors specific contracts and sends real-time notifications to a Telegram channel whenever a new event is registered.

### The Problem
Monitoring government contract changes is a tedious task. Important events (new links, documents, or status changes) can be easily missed if you don't check the portal constantly. I needed a lightweight service that tracks history and alerts stakeholders immediately.

### How it Works
The program follows a simple but effective logic to ensure no duplicates and monitoring:

1.  **State Management:** It uses a local `last_events.json` file to store the "last known state" (unique links, dates, and event text) for each contract.
2.  **Comparison Engine:** Every time the script runs, it fetches all current events from the source and compares them with the JSON file.
3.  **Telegram Integration:** If a new event is detected, it’s formatted and sent via the Telegram Bot API to a dedicated channel.
4.  **Persistence:** After notifying, the script updates the local JSON file to stay current for the next run.

### Setup & Installation
The project is configured for easy deployment via virtual environments and environment variables:

1. **Clone & Environment:**
   ```bash
   git clone git@github.com:EKuts3P/zakupki.git
   python -m venv venv
   source venv/bin/activate  # or venv\Scripts\activate on Windows
