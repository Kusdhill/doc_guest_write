import docx
import sys
import os
import re
import shutil


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
		print(doc_object.paragraphs[i].text)
		print(len(doc_object.paragraphs[i].runs))

		for run in doc_object.paragraphs[i].runs:
			if(run.bold):
				if(name_verified(run.text)):
					cleaned_name = clean_name(run.text)
					print("cleaned name = "+cleaned_name)
					names.append(cleaned_name)

	print(names)
	return names


# Verify that identified bold string is actually a name
def name_verified(text):
	print("IN NAME_VERIFIED")
	print(text)

	

	if(text=="" or len(text)<2):
		return False
	else:
		first_char = text[0]


	if(not first_char.isupper()):
		return False
	else:
		return True


# Cleans name of unnecessary bolded characters
def clean_name(text):
	print("cleaning "+text+"\n")
	last_char = text[-1]
	if(not last_char.isalpha()):
		return text[:-1]
	else:
		return text


# put name and text associated with name into dictionary
def copy_text(names, doc):
	name_with_text = {}

	#print(name_with_text)

	j = 0
	want_str = ""
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
			print("found name! "+names[j])
			
			if(j!=0):
				name_with_text[names[j-1]] = want_str
				#print("cleared want_str")
				want_str = ""
			if(j!=len(names)-1):
				j+=1
			
		if j==len(names)-1 and i==len(doc.paragraphs)-1:
			name_with_text[names[j]] = want_str
			want_str = ""

		want_str += text
		#print("added to "+want_str+"\n")

	for name in name_with_text:
		print name

	return name_with_text
		

# For each name, create a file, dump the text, and save the file
def dump_files(filename, names, copied):

	path = "./"+filename[0:-5]+"_created_files/"

	if os.path.exists(path):
		shutil.rmtree(path)

	os.makedirs(path)
	for i in range(0, len(names)):
		save_doc = docx.Document()
		save_doc.add_paragraph(copied[names[i]])
		save_doc.save(path+names[i]+".docx")


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
	names_with_text = copy_text(names, doc)
	print("creating files")
	dump_files(filename, names, names_with_text)
	


if __name__ == '__main__':
	main()