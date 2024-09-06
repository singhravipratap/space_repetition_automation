# Spaced Repetition System

This project is a Spaced Repetition System implemented using Python. It helps users to schedule and manage their study topics with spaced repetition intervals. 
The system uses an Excel file to store topics and their revision schedules, and it provides desktop notifications to remind users when it's time to revise a topic.

## Features

- Add new topics with spaced repetition schedules.
- Store and manage topics and their revision schedules in an Excel file.
- Display upcoming revisions in a GUI table.
- Send desktop notifications for due revisions.
- Schedule reminder checks to run every minute.

## Requirements

- Python 3.x
- `openpyxl` for handling Excel files
- `schedule` for scheduling tasks
- `tkinter` for the GUI
- `plyer` for desktop notifications

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/your-username/your-repository-name.git
    cd your-repository-name
    ```

2. Install the required Python packages:
    ```sh
    pip install openpyxl schedule plyer
    ```

## Usage

1. Run the script:
    ```sh
    python space_repetition_excel.py
    ```

2. Use the GUI to add new topics and view upcoming revisions.

## How It Works

- **Adding Topics**: Enter a topic in the GUI and click "Add Topic". The topic and its revision schedule will be added to the Excel file.
- **Checking Reminders**: The script checks for due revisions every minute and sends a desktop notification if a revision is due.
- **Displaying Revisions**: Click "Show Upcoming Revisions" in the GUI to display the list of topics and their revision schedules.

## File Structure

- `space_repetition_excel.py`: Main script containing the implementation of the spaced repetition system.
- `spaced_repetition_checklist_with_date_time.xlsx`: Excel file used to store topics and their revision schedules.

## Contributing

Contributions are welcome! Please create a pull request or open an issue to discuss any changes.

## License
This project is licensed under the MIT License.
