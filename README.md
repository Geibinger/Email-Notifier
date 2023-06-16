# Email Notification Script

This Python script reads an Excel file, checks whether the runout date is due, and if so, sends an email to the person using the SMTP client. 

## Installation

1. Clone this repository to your local machine.
2. Install [Python](https://www.python.org/downloads/) if not already installed.
3. Open a command prompt and navigate to the directory containing this script.
4. Create a virtual environment by running: `py -m venv venv`
5. Activate the virtual environment: `.\venv\Scripts\activate`
6. Install the required packages: `pip install -r requirements.txt`

## Configuration

Edit the `config.py` file in your text editor of choice to set the following variables:

- `username`: Your email username.
- `password`: Your email password.
- `smtp_server`: The SMTP server for your email provider.
- `smtp_port`: The SMTP port for your email provider.

## Running the Script

Ensure your virtual environment is activated (`.\venv\Scripts\activate`). Then, run the script with the command: `py notifier.py`.

Replace `notifier.py` with the name of the Python script.

## Automating the Script

To run this script automatically every day on a Windows system, you can use the Task Scheduler:

1. Open the Task Scheduler (you can search for it in the Start Menu).
2. Click on "Create Basic Task...".
3. Name the task and add a description, then click "Next".
4. Choose "Daily", then click "Next".
5. Set the start time and recurrence (1 day), then click "Next".
6. Choose "Start a program", then click "Next".
7. Browse to your Python script. In the "Start in (optional)" box, input the path to your Python script directory, then click "Next".
8. Click "Finish".

The script will now run automatically every day at the time you specified.
