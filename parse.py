import docx
import zipfile
import sys
import os
import re
import shutil
from docx.shared import Inches


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
		return verify_name(text[1:])
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
	print("\n\n\n\n\n\n"+"in copy_text---------------------------------------------------------------------"+"\n\n\n\n\n\n")
	name_with_text = {}
	text_list = []

	#print(names)

	j = 0
	all_found = False
	for i in range(0,len(doc.paragraphs)):
		line = doc.paragraphs[i]
		text = line.text
		bold = False
		next_name = False
		new_lines = 0
			
		# looking forward
		if(i<len(doc.paragraphs)-1):
			next_line = doc.paragraphs[i+1]
			next_text = next_line.text


		text_list.append(text)
		text+="*"

		print("doc = "+doc.paragraphs[i].text+"\n")

		# if all names have been found and parser is at the end of the doc
		# add text to the dictionary
		if(all_found and i==len(doc.paragraphs)-1):
			name_with_text[names[j]] = text_list
			#print("waow"+str(i))

		# if there is a bold run, set bold to true
		if(len(line.runs)>=2):
			print("\nfound runs")
			for k in range(0,len(line.runs)):
				print(line.runs[k].text)
				if(line.runs[k].bold):
					print("\nfound bold")
					bold = True

		if(j!=0 and re.search('[a-zA-Z]',text)==None):
			if(re.search('[a-zA-Z]',next_text)==None):
				pass
			else:
				print("append text_list = "+str(text_list))

				name_with_text[names[j-1]] = text_list
				#print(name_with_text[names[j-1]])
				#print("added, clear")
				text_list = []

		# if text is bold and it matches a name, increment j (pointer to lines)
		# if end of names list has been reached and name is found then set all_found to true
		if names[j] in text and bold:
			print("found name! "+names[j])
			if(j<len(names)-1):
				j+=1
			else:
				#print("found all names")
				all_found = True

	#print("text_list = "+str(text_list))
	return name_with_text


# get guest images from doc
def get_images(filename):

	image_locations = []
	stripped_filename = filename[0:-5]
	path = "./"+stripped_filename

	extract_directory = path+"_images"

	zip_ref = zipfile.ZipFile(filename, 'r')
	zip_ref.extractall(extract_directory)
	zip_ref.close()
	extract_path = extract_directory+"/word/media"

	for image in os.listdir(extract_path):
		image_locations.append(extract_path+"/"+image)

	image_locations.sort()

	return image_locations


# For each name, create a file, dump the text with images, and save the file
def dump_files(filename, names, copied, images):

	print(copied)

	asterisk = "*"
	path = "./"+filename[0:-5]+"_created_files/"
	all_guest_images = False

	if os.path.exists(path):
		shutil.rmtree(path)

	if(len(names)==len(images)):
		all_guest_images = True

	os.makedirs(path)
	for i in range(0, len(names)):
		entry = copied[names[i]]



		save_doc = docx.Document()
		

		for j in range(0,len(entry)):
			print(entry[j])

			

			if(entry[j]==""):
				print("nothing")

			if(j==0):
				para = save_doc.add_paragraph("")
				run = para.add_run(entry[j])
				run.bold = True
			else:
				if entry[j]!="":
					save_doc.add_paragraph(entry[j], style = 'List Bullet')

		if(all_guest_images):
			save_doc.add_picture(images[i],width=Inches(1.38), height=Inches(1.38))
		save_doc.save(path+names[i]+".docx")


# Clean created files
def clean_files(filename):
	stripped_filename = filename[0:-5]
	path = "./"+stripped_filename
	extract_directory = path+"_images"

	shutil.rmtree(extract_directory)


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
	print("getting images")
	guest_images = get_images(filename)
	print("creating files")
	dump_files(filename, names, names_with_text, guest_images)
	print("cleaning created files")
	clean_files(filename)
	

if __name__ == '__main__':
	main()