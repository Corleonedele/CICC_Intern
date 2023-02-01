import re

with open("test.txt") as f:
    for line in f:
        name = re.split("<|>", line)[2]

        print("<div><p>"+name+"</p><input id="+name+" type=\"text\" name="+name+" value=\"{{ input_"+name+" }}\"> </div> ")