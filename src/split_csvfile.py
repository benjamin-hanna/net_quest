import csv

class SplitFile():

    def __init__(self, file):
        self.file = file

    def split_file(self):
        file = self.file.rstrip()
        rest = []
        main = []

        with open(file, newline="") as csv_file:
            reader = csv.reader(csv_file, delimiter=',')

            main_header = next(reader)
            main_header_values = next(reader)
            zipped_main = list(zip(main_header, main_header_values))
            main.append(zipped_main)

            rest_header = next(reader)
            for row in reader:
                
                if row[1] == "TRUE": 
                    zipped_rest = list(zip(rest_header, row))
                    rest.append(zipped_rest)
        
        return {
            "main": main,
            "rest": rest
        }