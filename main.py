from pathlib import Path
from win32com.client import Dispatch

################################

# CHANGE THESE VARIABLES:
path = Path(r'path')  #path to the file you are looking for
fileName = "fileName.cpp" #name of the file you are looking for
groupName = "groupName"
emailAddress = "emailAddress"
firstName = "firstName"
lastName = "lastName"

################################

file = None

def findFile(path, fileName):
    for search in path.rglob(fileName):
        return search

    print("File not found")
    exit()

def modifyFile(file, path, groupName, emailAddress, firstName, lastName):
    extract = file.stem

    new_file = file.rename(path / f"identificationNumber.subjectName.{extract}.main.c")
    mailSubject = f"SUBJECT {groupName} {extract.upper()}"

    comments =(
    f"// {mailSubject}\n"
    f"// {firstName} {lastName}\n"
    f"// {emailAddress}\n"
    )

    oldContent = new_file.read_text()
    newContent = comments + oldContent
    new_file.write_text(newContent)

    return new_file, mailSubject

def prepareMail(file, mailSubject):
    outlook = Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = "mailRecipient@gmail.com"
    mail.Subject = mailSubject
    mail.Attachments.Add(str(file))
    mail.Display()

def main():
    file = findFile(path, fileName)
    file, mailSubject = modifyFile(file, path, groupName, emailAddress, firstName, lastName)
    prepareMail(file, mailSubject)

if __name__ == "__main__":
    main()