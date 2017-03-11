import os
from os import walk

path = "c:\\users\Tommy"
out = open("C:\\data\\find.txt","w") 
for (path, dirs, files) in os.walk(path):
    for file in files:
		fullname = path+file
		out.write(fullname + '\n')
out.close

