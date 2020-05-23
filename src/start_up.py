from pathlib import Path
import os
from os import path

class StartUp():

    def __init__(self, icon_path):
        self.icon_path = icon_path

    def start_up(self):

        try:
            folders = []
            base = Path(os.environ['HOME'] + '/Desktop/output/')
            folders.append(base)
            forms = Path(os.environ['HOME'] + '/Desktop/output/intake_forms/')
            folders.append(forms)
            emails = Path(os.environ['HOME'] + '/Desktop/output/emails/')
            folders.append(emails)

            for path in folders:

                if path.exists():
                    pass

                else: os.mkdir(path)

        except Exception as e: print(e)

        icon_path = self.icon_path
        
        with open(icon_path, 'r') as startup:
            print(startup.read())
            startup.close()

    def get_csv(self):
        csv = input("\nRemove all spaces from file name. \nDrag and drop .csv into terminal. \n")
        return csv


 
