import os
from pathlib import Path
from os.path import join
from render_text import render_text_templates, create_dictionaries
from render_gmail import render_gmail_templates

def text_or_gmail(prompt):
    while "Invalid answer":
        reply = str(input(prompt + '\n1: Write emails to text file \n'
                          '2: Send emails to gmail draft folder\n'
                          )).lower().strip()
        if reply[0] == '1':
            return True
        elif reply[0] == '2':
            return False
        else:
            print("Incorrect input. Please enter either 1 or 2.")

def input_workbook():

    imported_workbook = input("\nRemove all spaces from file name. \nDrag and drop .csv into terminal. \n")
    workbook = imported_workbook.rstrip()
    
    try:
        dictionary = create_dictionaries(workbook)
        return dictionary

    except:
        print("Error. Put opportunity workbook in input folder\n")
        print("Exiting...")

        exit(0)

def clean_directory(directory):
    
    try: 
        email_directory = [os.path.join(directory, f) for f in os.listdir(directory)]
        for file in email_directory:
            os.remove(file)
    except:
        print("Output directory not found in net_quest folder.")
        exit(0)

def check_directories():

    net_quest_output = Path(os.environ['HOME'] + '/Desktop/net_quest_output')
    net_quest_emails = Path(os.environ['HOME'] + '/Desktop/net_quest_output/emails')
    net_quest_intake_forms = Path(os.environ['HOME'] + '/Desktop/net_quest_output/intake_forms')

    if not os.path.exists(net_quest_output):
        os.makedirs(net_quest_output)

    if not os.path.exists(net_quest_emails):
        os.makedirs(net_quest_emails)
    clean_directory(net_quest_emails)

    if not os.path.exists(net_quest_intake_forms):
        os.makedirs(net_quest_intake_forms)
    clean_directory(net_quest_intake_forms)

def startup():

    with open(Path(os.environ['HOME'] + '/bin/net_quest/src/misc/startup'), 'r') as startup:
        print(startup.read())
        startup.close()

if __name__ == "__main__":

    dictionary = input_workbook()
    check_directories()
    startup()

    if text_or_gmail("Select from the following options: (Enter 1 or 2 into the terminal)") == True:
        render_text_templates(dictionary)
        exit(0)

    else:
        render_gmail_templates(dictionary)
        exit(0)
    