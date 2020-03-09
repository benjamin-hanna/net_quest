import csv
from jinja2 import Template, FileSystemLoader, Environment
import os
from pathlib import Path
from os.path import join
import delta_form_writer
import esi_form_writer
import jaa_form_writer
import sa_plus_form_writer

filtered_dictionaries = dict()
output_directory = Path(os.environ['HOME'] + '/Desktop/net_quest_output/emails/')

def split_file(file):

    lines = file.readlines()
    main = lines[:2]
    rest = lines[2:]

    return {
        "main": main,
        "rest": rest
    }

def string_to_bool(ord_dictionary):

    for item in ord_dictionary:
        if ord_dictionary[item] == 'TRUE':
            ord_dictionary[item] = True
        elif ord_dictionary[item] == 'FALSE':
            ord_dictionary[item] = False
        else:
            pass


def dash_to_whitespace(ord_dictionary):

    for item in ord_dictionary:
        if ord_dictionary[item] == '-':
            ord_dictionary[item] = ""
        else:
            pass


def update_global_dictionary_entries(csvlines):

    csv_reader = csv.DictReader(csvlines, delimiter=',')
    for row in csv_reader:
        filtered_dictionaries.update(row)


def return_list_of_ordered_dictionaries(csvlines):

    list_of_ordered_dictionaries = []
    csv_reader = csv.DictReader(csvlines, delimiter=',')
    for ord_dictionary in csv_reader:
        list_of_ordered_dictionaries.append(ord_dictionary)
    return list_of_ordered_dictionaries


def filter_networks(list_of_ordered_dictionaries):

    return list_of_ordered_dictionaries.get('is_make_template')


def create_dictionaries(path):

    with open(path, newline="") as csv_file:

        parts = split_file(csv_file)

        update_global_dictionary_entries(parts["main"])

        return_list_of_ordered_dictionaries(parts["rest"])

        filtered_dictionaries.update(
            {"network_info": [return_list_of_ordered_dictionaries(parts["rest"])]})

        for item in filtered_dictionaries['network_info']:
            for d in item:

                string_to_bool(d)
                dash_to_whitespace(d)

            filtered_list_of_dictionaries = [
                d for d in item if filter_networks(d)]

            filtered_dictionaries.update(
                {"network_info": [filtered_list_of_dictionaries]})

    return filtered_dictionaries


def list_of_networks(filtered_dictionary):

    list_of_networks = []

    for dictionary_list in filtered_dictionary['network_info']:
        for dictionary in dictionary_list:
            list_of_networks.append(dictionary['network_name'])
    return list_of_networks


def count_of_networks(list_of_networks):

    count_of_networks = len(list_of_networks)
    return count_of_networks


def render_text_templates(filtered_dictionary):

    global_items = list(filtered_dictionary.items())

    prospect = global_items[0][1]

    home = str(Path(os.environ['HOME']))

    list_of_templates = list_of_networks(filtered_dictionary)

    number_of_templates = count_of_networks(list_of_templates)

    print("\n", "number of emails: ", number_of_templates, "\n")

    loop_count = 0
    while loop_count < count_of_networks(list_of_networks(filtered_dictionary)):

        template_list = list_of_networks(filtered_dictionary)
        current_template = template_list[loop_count] + ".txt"

        for dictionary_list in filtered_dictionary['network_info']:

            env = Environment(loader=FileSystemLoader(
                searchpath=home + '/bin/net_quest/src/assets/templates'), trim_blocks=True, lstrip_blocks=True)


            template = env.get_template(current_template)

            final_email = template.render(

                # global variables
                prospect=global_items[0][1],
                rvp=global_items[1][1],
                hq=global_items[2][1],
                state=global_items[3][1],
                zip_code=global_items[4][1],
                city=global_items[5][1],
                start_date=global_items[6][1],
                due_date=global_items[7][1],
                broker=global_items[8][1],
                broker_address=global_items[9][1],
                broker_contact=global_items[10][1],
                broker_phone=global_items[11][1],
                broker_email=global_items[12][1],
                ees=global_items[13][1],
                ces=global_items[14][1],
                current_med_carrier=global_items[15][1],
                current_medical_funding=global_items[16][1],
                is_kaiser=global_items[17][1],
                current_pbm=global_items[18][1],
                current_dental_carrier=global_items[19][1],
                current_vision_carrier=global_items[20][1],
                current_stoploss=global_items[21][1],
                policy_period=global_items[22][1],
                specific_deductible=global_items[23][1],
                stoploss_aggregate_corridor=global_items[24][1],

                # instance variables
                network_name=dictionary_list[loop_count]['network_name'],
                is_inforce=dictionary_list[loop_count]['is_inforce'],
                is_quote=dictionary_list[loop_count]['is_quote'],
                is_discount_share=dictionary_list[loop_count]['is_discount_share'],
                is_geo_access=dictionary_list[loop_count]['is_geo_access'],
                is_disruption_analysis=dictionary_list[loop_count]['is_disruption_analysis'],
                is_claims_repricing=dictionary_list[loop_count]['is_claims_repricing'],
                is_zip_discount_analysis=dictionary_list[loop_count]['is_zip_discount_analysis'],
                is_transition_credits=dictionary_list[loop_count]['is_transition_credits'],
                is_questionnaire=dictionary_list[loop_count]['is_questionnaire'],
                is_stop_loss=dictionary_list[loop_count]['is_stop_loss'],
                is_benefit_deviation=dictionary_list[loop_count]['is_benefit_deviation']

            )
            # writes JAA intake form
            if dictionary_list[loop_count]['network_name'] == "anthem":
                print("\n", "writing JAA form", "\n")
                jaa_form_writer.write_jaa_intake_form(global_items)

            if dictionary_list[loop_count]['network_name'] == "empire":
                print("\n", "writing JAA form", "\n")
                jaa_form_writer.write_jaa_intake_form(global_items)

            if dictionary_list[loop_count]['network_name'] == "premera":
                print("\n", "writing JAA form", "\n")
                jaa_form_writer.write_jaa_intake_form(global_items)

            # writes SA+ intake form
            if dictionary_list[loop_count]['network_name'] == "blue_shield_ca":
                print("\n", "writing SA+ form" , "\n")
                sa_plus_form_writer.edit_sa_plus_form(global_items, filtered_dictionary)

            # writes Delta intake form
            if dictionary_list[loop_count]['network_name'] == "delta":
                print("\n", "writing Delta request form" , "\n")
                delta_form_writer.edit_delta_form(global_items)

            # writes ESI intake form
            if dictionary_list[loop_count]['network_name'] == "esi":
                print("\n", "writing ESI request form", "\n")
                esi_form_writer.edit_esi_form(global_items, filtered_dictionary)

            write_to_file(current_template, final_email)
            loop_count += 1

    rename_email(prospect)

def rename_email(prospect):

    email_subject_line_mapping = ["Blue Shield + Collective Health",
                                  "Anthem + Collective Health",
                                  "Empire + Collective Health",
                                  "Premera + Collective Health",
                                  "Cigna + Collective Health",
                                  "Delta + Collective Health",
                                  "Guardian + Collective Health",
                                  "CVS + Collective Health",
                                  "ESI + Collective Health",
                                  "VSP + Collective Health",
                                  "Eye Med + Collective Health",
                                  "Preliminary Stop Loss Quote"]

    for email in os.listdir(output_directory):

        try:
            txt_ending = ".txt"

            if email.endswith(txt_ending):
                if email.startswith("blue"):

                    email_subject_line = " ".join([email_subject_line_mapping[0], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("anthem"):

                    email_subject_line = " ".join([email_subject_line_mapping[1], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("empire"):

                    email_subject_line = " ".join([email_subject_line_mapping[2], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("premera"):

                    email_subject_line = " ".join([email_subject_line_mapping[3], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)


                elif email.startswith("cigna"):

                    email_subject_line = " ".join([email_subject_line_mapping[4], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("delta"):

                    email_subject_line = " ".join([email_subject_line_mapping[5], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("guardian"):

                    email_subject_line = " ".join([email_subject_line_mapping[6], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("cvs"):

                    email_subject_line = " ".join([email_subject_line_mapping[7], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("esi"):

                    email_subject_line = " ".join([email_subject_line_mapping[8], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("vsp"):

                    email_subject_line = " ".join([email_subject_line_mapping[9], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("eye_med"):

                    email_subject_line = " ".join([email_subject_line_mapping[10], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("sun_life.txt"):

                    email_subject_line = " ".join([email_subject_line_mapping[11], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("berkley"):

                    email_subject_line = " ".join([email_subject_line_mapping[11], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("iisi"):

                    email_subject_line = " ".join([email_subject_line_mapping[11], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("tokio_marine"):

                    email_subject_line = " ".join([email_subject_line_mapping[11], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("symetra"):

                    email_subject_line = " ".join([email_subject_line_mapping[11], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("voya"):

                    email_subject_line = " ".join([email_subject_line_mapping[11], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

                elif email.startswith("optum"):

                    email_subject_line = " ".join([email_subject_line_mapping[11], "-" , prospect])
                    path_email_subject_line = join(output_directory, email_subject_line)
                    path_email = join(output_directory, email)
                    os.rename(path_email, path_email_subject_line)
                    print("\n", "email written: ", email_subject_line)

        except:
            print("\n", "Emails printed to the output folder could not be renamed.", "\n")
            exit(0)

def write_to_file(template_to_use, final_email):

    file_path = output_directory.joinpath(template_to_use)

    with open(file_path, 'w') as file:
        file.write(final_email)
