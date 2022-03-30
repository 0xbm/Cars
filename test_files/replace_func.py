'''
#input file
fin = open("text_to_replace.py", "rt")
#output file to write the result to
fout = open("text_to_replace.py", "wt")
#for each line in the input file
for line in fin:
	#read replace the string and write to output file
	fout.write(line.replace('wb', 'wb1'))
#close input and output files
fin.close()
fout.close()
'''


#read input file
fin = open("text_to_replace.py", "rt")
#read file contents to string
data = fin.read()
#replace all occurrences of the required string
data = data.replace('wb2', 'wb3')
#close the input file
fin.close()
#open the input file in write mode
fin = open("text_to_replace.py", "wt")
#overrite the input file with the resulting data
fin.write(data)
#close the file
fin.close()