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

	# How to identify a name?
	# Every name is in bold
	# Most names are followed by a comma, but not all
	# Newline (empty paragraph) preceeds every name
	# Line succeeding every name begins with "Affiliation"


	##
	##for para in doc_object.paragraphs:
	##	print para.text +"\\n"

	#print("\n"+doc_object.paragraphs[1].runs[1].text)


	full_text = ""
	for para in doc_object.paragraphs:
		full_text += para.text

	names = []
	for i in range(0, len(doc_object.paragraphs)):
		if(len(doc_object.paragraphs[i].runs)>=2):
			run = doc_object.paragraphs[i].runs[1]
			if(run.bold):
				names.append(run.text)

	print(names)




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
	find_names(doc)


if __name__ == '__main__':
	main()