import os

class WriteObject():

    def __init__(self, path, email, network_name):
        self.path = path
        self.email = email
        self.network_name = network_name
    
    def write_object(self):
        path = self.path
        email = self.email
        network_name = self.network_name

        file_path = os.path.join(path, network_name)

        with open(file_path, "w") as file:
            file.write(email)



