College Lab Assignment Automation Script
This repository contains a Python script that automates the process of preparing and sending lab assignments to the anti-plagiarism system at my college

THE SCRIPT
- finds the specified source file in the project directory
- renames it according to the required university format
- inserts the required header comments (subject, author, email)
- opens a ready-to-send Outlook email with the file attached

!!! This project intentionally uses placeholder values instead of real personal data

FEATURES
- Searches recursively for your `.c` or `.cpp` source file
- Applies the required pattern: `identityNumber.subjectName.groupName.main.c`.
- Adds the first three required comment lines:
  1. Mail subject  
  2. Author name  
  3. Email
- Creates a ready-to-send email with the attachment included.

HOW IT WORKS
1. Update the variables at the top of `main.py` (path, filenames, ...).
2. Run the script.
3. Your file will be found, renamed, modified, and attached to an Outlook draft email.

REQUIREMENTS
- Python 3.x 
- Outlook installed and configured
- `pywin32` package:
  bash: `pip install pywin32`
