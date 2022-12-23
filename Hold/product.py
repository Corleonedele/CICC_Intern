
import pickle

file_name = "code.pkl"

class Product():
    def __init__(self, name, code_SZ, code_Mail, code_BJ) -> None:
        self.name = name
        self.code_SZ = code_SZ
        self.code_Mail = code_Mail
        self.code_BJ = code_BJ
    




    
test = Product("1", "2", "3", "4")
print(test.code_BJ)
with open(file_name, 'wb') as translate_file:
    pickle.dump(test, translate_file)

with open(file_name, 'rb') as translate_file:
    x = pickle.load(translate_file)

