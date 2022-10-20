# Python3 program to convert docx to pdf
# using docx2pdf module
 
# Import the convert method from the
# docx2pdf module
from docx2pdf import convert
 
# Converting docx present in the same folder
# as the python file
# convert("RESULT CANCELLATION CERTIFICATE.docx")
 
# Converting docx specifying both the input
# and output paths
convert("MIGRATION CERTIFIACTE.docx", "MIGRATION CERTIFIACTE.pdf")

 

 
