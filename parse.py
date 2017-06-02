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
		complete_run = ""
		#print(doc_object.paragraphs[i].text)
		#print(len(doc_object.paragraphs[i].runs))

		for run in doc_object.paragraphs[i].runs:
			# need to get all of bolded line instead of fragmented lines
			if(run.bold):
				complete_run += run.text

		if(verify_name(complete_run)):
			cleaned_name = clean_name(complete_run)
			print("cleaned name = "+cleaned_name+"\n")
			names.append(cleaned_name)

	print(names)

	mend(names)

	return names


# Verify that identified bold string is actually a name
def verify_name(text):
	print("IN VERIFY_NAME: "+text)
	
	if(not (" " in text)):
		print("no space, returning false")
		return False
	if(text=="" or len(text)<2):
		print("empty/small text, returning false")
		return False
	else:
		first_char = text[0]
		last_char = text[-1]
	if(not first_char.isupper()):
		print("first char not upper, returning false")
		return False
	if(not (last_char.isalpha() or last_char!="," or last_char!=":" or last_char!=" " or last_char!="\"")):
		print("last char not alpha or , or : or space or \", returning false")
		return False
	else:
		return True


# Checks for fragmented names
def mend(names):
	print("in mend")
	for i in range(0,len(names)-2):
		if " " not in names[i]:
			join_indeces(names,i,i+1)


# Joins fragmented names
def join_indeces(list_n, left, right):
	first_name = list_n[left]
	last_name  = list_n[right]
	name_string = first_name+" "+last_name

	del list_n[left]
	del list_n[left]
	list_n.insert(left, name_string)
	print("list_n = "+str(list_n))


# Cleans name of unnecessary bolded characters
def clean_name(text):
	print("cleaning "+text+"\n")
	last_char = text[-1]
	print("last_char = "+last_char)
	print("is alpha? "+str(last_char.isalpha()))
	#print(text[0:-1])
	if(not last_char.isalpha()):
		#print("bad return?")
		print("re-cleaning "+text[0:-1])
		return clean_name(text[0:-1])
	else:
		print("returning "+text)
		return text


# put name and text associated with name into dictionary
def copy_text(names, doc):
	name_with_text = {}

	#print(names)

	j = 0
	all_found = False
	want_str = ""
	for i in range(0,len(doc.paragraphs)):
		line = doc.paragraphs[i]

		text = line.text
		bold = False
		next_name = False
		new_lines = 0
			
		#print(i)
		if(i<len(doc.paragraphs)-1):
			next_line = doc.paragraphs[i+1]
			next_text = next_line.text

		want_str+=text
		#print("text = "+text)
		#print(want_str+"\n\n\n")

		#print("doc = "+doc.paragraphs[i].text+"\n")


		if(all_found and i==len(doc.paragraphs)-1):
			name_with_text[names[j]] = want_str
			#print("waow"+str(i))


		if(len(line.runs)>=2):
			#print("\nfound runs")
			for k in range(0,len(line.runs)):
				if(line.runs[k].bold):
					#print("\nfound bold")
					bold = True

		if(j!=0 and re.search('[a-zA-Z]',text)==None):
			if(re.search('[a-zA-Z]',next_text)==None):
				pass
			else:
				name_with_text[names[j-1]] = want_str
				#print(name_with_text[names[j-1]])
				#print("added, clear")
				want_str = ""

		#print("\nlooking for "+names[j])
		if names[j] in text and bold:
			#print("found name! "+names[j])
			if(j<len(names)-1):
				j+=1
			else:
				#print("found all names")
				all_found = True


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