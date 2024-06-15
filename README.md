# Attendance Tracking System

This Attendance Tracking System is a Python-based project that tracks students' attendance and sends warning emails to students and staff when attendance falls below a certain threshold. The project uses Google Sheets for data storage and the Gmail API for sending emails.

## Features

- Track attendance for multiple subjects.
- Send warning emails to students with low attendance.
- Notify staff members about students with low attendance.
- Use environment variables for secure management of sensitive information.

## Prerequisites

- Python 3.x
- Google Account with access to the Gmail API
- Google Sheets with student attendance data
- The following Python packages:
  - `openpyxl`
  - `google-auth`
  - `google-auth-oauthlib`
  - `google-auth-httplib2`
  - `google-api-python-client`
  - `python-dotenv`

## Setup

### Step 1: Clone the Repository

```sh
git clone https://github.com/your-username/Attendance_Tracking_System.git
cd Attendance_Tracking_System
```

### Step 2: Install the Required Libraries

```sh
pip install openpyxl google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client python-dotenv
```

### Step 3: Add Your Google API Credentials

- Place your `client_id.json` file in the project directory. This file contains your Google OAuth Client ID and Client Secret.
- Create a `.env` file in the project directory with the following content:

```env
CLIENT_SECRET_FILE=client_id.json
```

### Step 4: Prepare the Google Sheets

Ensure you have a Google Sheets file with student attendance data. The file should have the following structure:

- The first sheet named `Sheet1`.
- The first row should contain column headers: Roll Number, Email, DBMS Attendance, DS Attendance, Python Attendance.
- Each subsequent row should contain attendance data for a student.

For example:

| Roll Number | Email                | DBMS Attendance | DS Attendance | Python Attendance |
|-------------|----------------------|-----------------|---------------|-------------------|
| 1           | student1@example.com | 0               | 0             | 0                 |
| 2           | student2@example.com | 1               | 0             | 2                 |
| 3           | student3@example.com | 2               | 1             | 3                 |

### Step 5: Running the Script

```sh
python att_track_sys.py
```

Follow the prompts to enter the subject and the number of absentees.

## Detailed Instructions

### Using the Script

1. **Enter the Subject:**

   When prompted, enter the subject number:
   - `1` for DBMS
   - `2` for DS
   - `3` for Python

2. **Enter the Number of Absentees:**

   Enter the number of students who were absent.

3. **Enter the Roll Numbers of Absentees:**

   Enter the roll numbers of the absent students.

4. **Send Emails:**

   The script will send warning emails to the absent students and notify the staff member.

## Project Structure

- `att_track_sys.py`: The main script to track attendance and send emails.
- `client_id.json`: The Google API credentials file (not included in the repository for security reasons).
- `token.json`: The token file for the Gmail API (generated automatically after the first run).
- `.env`: Environment file containing the path to the `client_id.json` file.
- `attendance.xlsx`: The Excel file with student attendance data.

## Functions

- **`savefile()`:** Save the updated attendance data to the Excel file.
- **`get_service()`:** Authenticate and return the Gmail API service.
- **`send_email(to, subject, body)`:** Send an email using the Gmail API.
- **`mailstu(l1, m)`:** Send warning emails to a list of students.
- **`mailstaff()`:** Send notification email to the staff member.
- **`check(no_of_days, row_num, b)`:** Check attendance and send appropriate emails.

### Sample Code Snippet

```python
def mailstu(l1, m):
    for student_email in l1:
        send_email(student_email, 'Attendance Warning', m)

def mailstaff(l2, l3, subject):
    if l2:
        msg1 = "You have lack of attendance in " + subject + "!!!"
        msg2 = "The following students have lack of attendance in your subject: " + l2
        for email in l3:
            send_email(email, 'Attendance Alert', msg1)
        send_email(staff_mails[subject_to_index(subject)], 'Lack of Attendance Report', msg2)
```

## Contributing

Contributions are welcome! Please create an issue or submit a pull request.

## License

This project is licensed under the MIT License.

## Contact

For any questions or feedback, please contact [your-email@example.com].

## Troubleshooting

If you encounter issues pushing to GitHub, ensure no sensitive information (like `client_id.json` or `token.json`) is included in your commits. You can add these files to your `.gitignore` to prevent accidental pushes.

### Adding Files to `.gitignore`

Create a `.gitignore` file in your project directory with the following content:

```plaintext
client_id.json
token.json
.env
```

This will prevent Git from tracking these files.

### Removing Sensitive Data from History

If you have already committed sensitive files, you need to remove them from the history. Here is an example using the `bfg` tool:

1. Install `bfg`:

   ```sh
   brew install bfg
   ```

2. Run `bfg` to remove sensitive files:

   ```sh
   bfg --delete-files client_id.json
   bfg --delete-files token.json
   ```

3. Clean up the repository:

   ```sh
   git reflog expire --expire=now --all
   git gc --prune=now --aggressive
   ```

4. Force push to GitHub:

   ```sh
   git push --force
   ```

## Acknowledgements

- [Google API Python Client](https://github.com/googleapis/google-api-python-client)
- [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [Dotenv](https://pypi.org/project/python-dotenv/)

---
```

Make sure to update `"your-username/Attendance_Tracking_System"` and `"your-email@example.com"` with your actual GitHub username and email address. This `README.md` file provides comprehensive instructions and details for setting up, using, and contributing to your Attendance Tracking System project.
