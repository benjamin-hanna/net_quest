from jinja2 import Template, FileSystemLoader, Environment

class TemplateObject():

    def __init__(self, main, rest, loop_count):
        self.main = main
        self.rest = rest
        self.loop_count = loop_count
        
    
    def template_object(self):
        main = self.main
        rest = self.rest
        i = self.loop_count
        network_name = rest[i][0][1]
        env = Environment(loader=FileSystemLoader(searchpath="./src/assets/templates/"), trim_blocks=True, lstrip_blocks=True)
        template = env.get_template(network_name + ".txt")
        email = template.render(
                prospect = main[0][0][1],
                rvp = main[0][1][1],
                hq = main[0][2][1],
                state = main[0][3][1],
                zip_code = main[0][4][1],
                city = main[0][5][1],
                start_date = main[0][6][1],
                due_date = main[0][7][1],
                broker = main[0][8][1],
                broker_address = main[0][9][1],
                broker_contact = main[0][10][1],
                broker_phone = main[0][11][1],
                broker_email = main[0][12][1],
                ees = main[0][13][1],
                ces = main[0][14][1],
                current_med_carrier = main[0][15][1],
                current_medical_funding = main[0][16][1],
                is_kaiser = main[0][17][1],
                current_pbm = main[0][18][1],
                current_dental_carrier = main[0][19][1],
                current_vision_carrier = main[0][20][1],
                current_stoploss = main[0][21][1],
                policy_period = main[0][22][1],
                specific_deductible = main[0][23][1],
                stoploss_aggregate_corridor = main[0][24][1],

                network_name = network_name,
                is_inforce = rest[i][2][1],
                is_quote = rest[i][3][1],
                is_discount_share = rest[i][4][1],
                is_geo_access = rest[i][5][1],
                is_disruption_analysis = rest[i][6][1],
                is_claims_repricing = rest[i][7][1],
                is_zip_discount_analysis = rest[i][8][1],
                is_network_guarantee = rest[i][9][1],
                is_transition_credits = rest[i][10][1],
                is_questionnaire = rest[i][11][1],
                is_stop_loss = rest[i][12][1],
                is_benefit_deviation = rest[i][13][1]

            )
        return {
            "network_name": network_name,
            "email": email
        }
