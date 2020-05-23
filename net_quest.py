from src.split_csvfile import SplitFile
from src.start_up import StartUp
from src.template_object import TemplateObject
from src.write_object import WriteObject
from src.form_object import FormObject
from pathlib import Path
import os

# startup
start = StartUp("./src/assets/static/start_up")
start.start_up()
csv = start.get_csv()

# splits the file into main and rest
file = SplitFile(csv)
split_file = file.split_file()
main = split_file.get("main")
rest = split_file.get("rest")

print("\n")
# creates templates from split file
for i in range(len(rest)):
    # creates email object
    template_object = TemplateObject(main, rest, i)
    email_template = template_object.template_object()
    network_email = email_template.get("email")
    network_name = email_template.get("network_name")

    # writes email objects to file
    email = WriteObject(Path(os.environ['HOME'] + '/Desktop/output/emails'), network_email, network_name)
    email.write_object()

    # writes form objects to file
    form = FormObject(Path(os.environ['HOME'] + '/Desktop/output/intake_forms'), network_name, main, rest, i)
    form.network_form_switcher()

    # iterator
    print("Email written for: ", network_name)
    i = i + 1

print("\nExiting net_quest...\n")
exit()