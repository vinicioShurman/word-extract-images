import docx2txt

input_loc = input("Your docx location: ")
output_loc = input("Output location: ")

text = docx2txt.process(input_loc, output_loc)

# Don't forget to 