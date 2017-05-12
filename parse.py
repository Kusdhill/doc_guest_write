import docx
import sys
import os

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



def main():
	print("checking command line arguments")
	if len(sys.argv)!=2:
		sys.exit("usage: python compare.py filename.docx")
	else:
		print("verifying file extension")
		check_extension(sys.argv[1])
		print("verifying existence")
		check_existence(sys.argv[1])

	filename = sys.argv[1]

	doc = docx.Document(filename)


	for pars in doc.paragraphs:
		print pars.text +"\\n"


if __name__ == '__main__':
	main()