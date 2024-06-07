# Email Analysis Script

This script connects to an IMAP email server, retrieves emails based on specific criteria, filters them, and generates an Excel report with a chart showing the number of emails received per day.

## Prerequisites

1. **Python 3.** If you don't have Python installed, follow the instructions on the official Python website to install it: [Install Python](https://www.python.org/downloads/)

2. **IMAP email server credentials.** You need access to an IMAP email server and the necessary credentials (username and password).

## Setup Instructions
### Step 1: Clone the Repository
Clone this repository to your local machine.

```sh
git clone <repository-url>
cd <repository-directory>
```

### Step 2: Create a Virtual Environment
Create a virtual environment to manage dependencies (inside the repository folder).

```
python -m venv myenv-Emails
```

### Step 3: Activate the Virtual Environment
Activate the virtual environment. The command differs depending on your operating system:

* For Windows:

```
myenv-Emails\Scripts\activate
```

* For macOS and Linux:
```
source myenv-Emails/bin/activate
```

### Step 4: Install Required Packages
Install the required packages using `pip`.

```
pip install imaplib2 pandas matplotlib openpyxl configparser
```

### Step 5: Configure Your Email Credentials
Create a config.ini file in the same directory as the script and add your email credentials and IMAP server details.

```
# config.ini

[email]
username = your_email@example.com
password = your_password
imap_server = imap.example.com
imap_port = 993
```
Don't use any quotes to surround the values in `config.ini`.

### Step 6: Run the Script
Run the script to process the emails and generate the Excel report.

```
python email_analysis.py
```

### Step 7: Check the Output
The script will generate an email_analysis.xlsx file in the same directory. This file contains the total email counts and individual folder counts, along with a chart showing the number of emails received per day.

### Step 8: Deactivate the virtual environment
This is optional step   to deactivate the virtual environment, after finishing the work with the script.
Just type in the console from where you ran the script following command:
`deactivate`

## Dependencies
* `imaplib2`: For connecting to the IMAP email server.
* `pandas`: For data manipulation and analysis.
* `matplotlib`: For plotting charts.
* `openpyxl`: For creating and modifying Excel files.
* `configparser`: For reading configuration files.

These dependency libraries can be installed by:
`pip install imaplib2 pandas matplotlib openpyxl configparser`


## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.