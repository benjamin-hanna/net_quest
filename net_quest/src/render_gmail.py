from jinja2 import Template, FileSystemLoader, Environment
import gmail_api
import os
from pathlib import Path
import jaa_form_writer
import sa_plus_form_writer
import delta_form_writer
import esi_form_writer

output_directory = Path('../output')


def clean_directory():

    try:

        files_to_remove = [os.path.join(output_directory, f)
                           for f in os.listdir(output_directory)]
        for file in files_to_remove:
            os.remove(file)

    except:
        print("output directory not found in net_quest folder")


def list_of_networks(filtered_dictionary):

    list_of_networks = []

    for dictionary_list in filtered_dictionary['network_info']:
        for dictionary in dictionary_list:
            list_of_networks.append(dictionary['network_name'])
    return list_of_networks


# produces a count for the while loop to cycle through
def count_of_networks(list_of_networks):

    count_of_networks = len(list_of_networks)
    return count_of_networks


def render_gmail_templates(filtered_dictionary):

    clean_directory()

    # creates list for global variables to call
    global_items = list(filtered_dictionary.items())

    prospect = global_items[0][1]

    list_of_templates = list_of_networks(filtered_dictionary)

    number_of_templates = count_of_networks(list_of_templates)

    print("number of templates: ", number_of_templates, "\n")

    loop_count = 0
    while loop_count < count_of_networks(list_of_networks(filtered_dictionary)):

        template_list = list_of_networks(filtered_dictionary)
        current_template = template_list[loop_count] + ".txt"

        for dictionary_list in filtered_dictionary['network_info']:

            env = Environment(trim_blocks=True, lstrip_blocks=True, loader=FileSystemLoader(
                searchpath="./assets/templates"))

            email_document = env.get_template(current_template)

            final_email = email_document.render(

                # global variables
                prospect=global_items[0][1],
                rvp=global_items[1][1],
                broker=global_items[2][1],
                broker_address=global_items[3][1],
                broker_contact=global_items[4][1],
                broker_phone=global_items[5][1],
                broker_email=global_items[6][1],
                start_date=global_items[7][1],
                ees=global_items[8][1],
                medical_covered_ees=global_items[9][1],
                dental_covered_ees=global_items[10][1],
                vision_covered_ees=global_items[11][1],
                current_medical_funding=global_items[12][1],
                current_med_carrier=global_items[13][1],
                current_dental_carrier=global_items[14][1],
                current_vision_carrier=global_items[15][1],
                current_pbm=global_items[16][1],
                current_stoploss=global_items[17][1],
                hq=global_items[18][1],
                is_kaiser=global_items[19][1],
                due_date=global_items[20][1],
                commission=global_items[21][1],
                policy_period=global_items[22][1],
                specific_deductible=global_items[23][1],
                agg_deductible_attachment_point=global_items[24][1],

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
                print("\n", "writing SA+ form", "\n")
                sa_plus_form_writer.edit_sa_plus_form(
                    global_items, filtered_dictionary)

            # writes Delta intake form
            if dictionary_list[loop_count]['network_name'] == "delta":
                print("\n", "writing Delta request form", "\n")
                delta_form_writer.edit_delta_form(global_items)

            # writes ESI intake form
            if dictionary_list[loop_count]['network_name'] == "esi":
                print("writing ESI request form", "\n")
                esi_form_writer.edit_esi_form(
                    global_items, filtered_dictionary)

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

            for subject_line in email_subject_line_mapping:

                try:

                    if current_template.startswith("blue"):
                        current_template = " ".join(
                            [email_subject_line_mapping[0] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("anthem"):
                        current_template = " ".join(
                            [email_subject_line_mapping[1] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("empire"):
                        current_template = " ".join(
                            [email_subject_line_mapping[2] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("premera"):
                        current_template = " ".join(
                            [email_subject_line_mapping[3] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("cigna"):
                        current_template = " ".join(
                            [email_subject_line_mapping[4] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("delta"):
                        current_template = " ".join(
                            [email_subject_line_mapping[5] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("guardian"):
                        current_template = " ".join(
                            [email_subject_line_mapping[6] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("cvs"):
                        current_template = " ".join(
                            [email_subject_line_mapping[7] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("esi"):
                        current_template = " ".join(
                            [email_subject_line_mapping[8] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("vsp"):
                        current_template = " ".join(
                            [email_subject_line_mapping[9] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("eye_med"):
                        current_template = " ".join(
                            [email_subject_line_mapping[10] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("sun_life.txt"):
                        current_template = " ".join(
                            [email_subject_line_mapping[11] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("berkley"):
                        current_template = " ".join(
                            [email_subject_line_mapping[11] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("iisi"):
                        current_template = " ".join(
                            [email_subject_line_mapping[11] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("tokio_marine"):
                        current_template = " ".join(
                            [email_subject_line_mapping[11] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("symetra"):
                        current_template = " ".join(
                            [email_subject_line_mapping[11] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("voya"):
                        current_template = " ".join(
                            [email_subject_line_mapping[11] + ": " + prospect + " " + "(01/01/20)"])

                    elif current_template.startswith("optum"):
                        current_template = " ".join(
                            [email_subject_line_mapping[11] + ": " + prospect + " " + "(01/01/20)"])

                    else:
                        pass

                except:
                    raise

            gmail_api.make_drafts(current_template, final_email)
            loop_count += 1
