import docx
import sys
import os
import re

# verifies that a given filename has .docx extension
# if it does not, an error is thrown
def check_extension(filename):
	extension_flag = 0
	extension = ""
	good_extension = "docx"

	for char in filename:
		if extension_flag==0:
			if char==".":
				extension_flag = 1
		else:
			extension+=char

	if extension!=good_extension:
		sys.exit("must pass in .docx files")


# verifies that file exists
# if it does not, an error is thrown
def check_existence(filename):
	if not os.path.isfile(filename):
		sys.exit("file must exist")


# parses file for names of guests
def find_names(doc_object):
	names = []
	for i in range(0, len(doc_object.paragraphs)):
		if(len(doc_object.paragraphs[i].runs)>=2):
			run = doc_object.paragraphs[i].runs[1]
			if(run.bold):
				names.append(run.text)

	#print(names)
	return names

# put name and text associated with name into dictionary
def copy_text(names, doc):
	name_with_text = {}

	#print(name_with_text)

	j = 0
	want_str = ""
	want = []
	for i in range(0,len(doc.paragraphs)):
		line = doc.paragraphs[i]
		text = line.text
		bold = False

		#print(text)

		if(len(line.runs)>=2):
			if(line.runs[1].bold):
				bold = True
			else:
				bold = False

		if names[j] in text and bold:
			#print(j)
			#print("found name! "+names[j])
			
			if(j!=0):
				want.append(want_str)
				name_with_text[names[j-1]] = want_str
				#print("cleared want_str")
				want_str = ""
			if(j!=len(names)-1):
				j+=1
			
		if j==len(names)-1 and i==len(doc.paragraphs)-1:
			want.append(want_str)
			name_with_text[names[j]] = want_str
			want_str = ""

		want_str += text
		#print("added to "+want_str+"\n")

	return name_with_text
		

def main():
	print("checking command line arguments")
	if len(sys.argv)!=2:
		sys.exit("usage: python parse.py filename.docx")
	else:
		print("verifying file extension")
		check_extension(sys.argv[1])
		print("verifying existence")
		check_existence(sys.argv[1])

	filename = sys.argv[1]

	doc = docx.Document(filename)
	print("finding names")
	names = find_names(doc)
	print("copying text")
	text = copy_text(names, doc)
	print("creating files")
	


if __name__ == '__main__':
	main()