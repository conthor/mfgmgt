import os
from os import walk

path = "c:\\users\Tommy"
listf="C:\\data\\find.txt"
out = open(listf,"w") 
for (path, dirs, files) in os.walk(path):
    for file in files:
		fullname = path+file
		out.write(fullname + '\n')
out.close
