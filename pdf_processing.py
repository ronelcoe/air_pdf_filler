import os
import re

import fitz  # requires fitz, PyMuPDF
import pdfrw
import subprocess
import os.path
import sys
from datetime import datetime, timedelta
import pandas as pd

class ProcessPdf:

    def __init__(self, temp_directory, output_file):
        print('\n##########| Initiating Pdf Creation Process |#########\n')
        
        print('\nDirectory for storing all temporary files is: ', temp_directory)
        self.temp_directory = temp_directory
        print("Final Pdf name will be: ", output_file)
        self.output_file = output_file

    def add_data_to_pdf(self, template_path, data):
        template = pdfrw.PdfReader(template_path)

        finalOutputDirectory = self.temp_directory + data['checkin'].replace(r"/","")
        finalOutputFile = finalOutputDirectory + "/" + self.output_file
        
        if not os.path.isfile(finalOutputFile):
            if not os.path.exists(finalOutputDirectory):
                os.makedirs(finalOutputDirectory)
        
            for page in template.pages:
                annotations = page['/Annots']
                if annotations is None:
                    continue

                for annotation in annotations:
                    if annotation['/T']:
                        key = annotation['/T'][1:-1]
                        if re.search(r'.-[0-9]+', key):
                            key = key[:-2]

                        if key in data:
                            annotation.update(
                                pdfrw.PdfDict(Contents='{}'.format(data[key]))
                            )

                            annotation.update(
                                    pdfrw.PdfDict(V=self.encode_pdf_string(data[key], 'string'))
                            )
                            annotation.update(pdfrw.PdfDict(Ff=1))

            template.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))

            pdfrw.PdfWriter().write(finalOutputFile, template)  

    def encode_pdf_string(self, value, type):
        if type == 'string':
            if value:
                return pdfrw.objects.pdfstring.PdfString.encode(value.upper())
            else:
                return pdfrw.objects.pdfstring.PdfString.encode('')
        elif type == 'checkbox':
            if value == 'True' or value == True:
                return pdfrw.objects.pdfname.BasePdfName('/Yes')
                # return pdfrw.objects.pdfstring.PdfString.encode('Y')
            else:
                return pdfrw.objects.pdfname.BasePdfName('/No')
                # return pdfrw.objects.pdfstring.PdfString.encode('')
        return ''

    def delete_temp_files(self, pdf_list):
        print('\nDeleting Temporary Files...')
        for path in pdf_list:
            try:
                os.remove(path)
            except:
                pass

    def compress_pdf(self, input_file_path, power=3):
        """Function to compress PDF via Ghostscript command line interface"""
        quality = {
            0: '/default',
            1: '/prepress',
            2: '/printer',
            3: '/ebook',
            4: '/screen'
        }

        output_file_path = self.temp_directory + 'compressed.pdf'

        if not os.path.isfile(input_file_path):
            print("\nError: invalid path for input PDF file")
            sys.exit(1)

        if input_file_path.split('.')[-1].lower() != 'pdf':
            print("\nError: input file is not a PDF")
            sys.exit(1)

        print("\nCompressing PDF...")
        initial_size = os.path.getsize(input_file_path)
        subprocess.call(['gs', '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4',
                        '-dPDFSETTINGS={}'.format(quality[power]),
                        '-dNOPAUSE', '-dQUIET', '-dBATCH',
                        '-sOutputFile={}'.format(output_file_path),
                         input_file_path]
        )
        final_size = os.path.getsize(output_file_path)
        ratio = 1 - (final_size / initial_size)
        print("\nCompression by {0:.0%}.".format(ratio))
        print("Final file size is {0:.1f}MB".format(final_size / 1000000))
        return output_file_path

    def date_extention(self, number):
        if number%10 == 1:
            return '%dst' % number
        if number%10 == 2:
            return '%dnd' % number
        if number%10 == 3:
            return '%drd' % number
        if (number%10 >= 4) or (number%10== 0):
            return '%dth' % number

    def massageCSVData(self, value):
        return value.replace(r" \(.*\)","")

BASE_DIRECTORY =  os.getcwd()
TEMPLATE_PATH = BASE_DIRECTORY + '/templates/'
FILLED_UP_FORMS = BASE_DIRECTORY + '/filled_out_forms/'
AUTH_LETTER_TEMPLATE_FILENAME = 'AIR_Guest_Template.pdf'
HEALTH_DECLARATION_TEMPLATE_FILENAME = 'AIR_Health_Declaration.pdf'
AUTH_LETTER_OUTPUT_FILE_NAME = 'Autorization_'
HEALTH_DECLARATION_OUTPUT_FILE_NAME = 'Health_Declaration_'
CSV_RESERVATION = BASE_DIRECTORY + '/reservations.csv'
GATEPASS_CLEANING_TEMPLATE_FILENAME = 'gate_pass_cleaning_template.pdf'
WORK_PERMIT_CLEANING_TEMPLATE_FILENAME = 'work_permit_cleaning_template.pdf'
GATEPASS_CLEANING_OUTPUT_FILE_NAME = 'gatepass_cleaning_'
WORK_PERMIT_CLEANING_OUTPUT_FILE_NAME = 'workpermit_cleaning_'


csvHeader = ['Guest name', 'Start date', 'End date', 'Confirmation code']
data = pd.read_csv(CSV_RESERVATION, skipinitialspace=True, usecols=csvHeader)
df = pd.DataFrame(data)

# currentDate = datetime.today().strftime("%d")
# currentMonth = datetime.today().strftime("%B")
# currentFullYear = datetime.today().strftime("%Y")
# currentShortYear = datetime.today().strftime("%y")

for index, guest in df.iterrows():
    guestName = guest[csvHeader[0]].replace(r" \(.*\)","")
    checkin = guest[csvHeader[1]].replace(r" \(.*\)","")
    checkout = guest[csvHeader[2]].replace(r" \(.*\)","")
    confirmationCode = guest[csvHeader[3]].replace(r" \(.*\)","")

    pdf = ProcessPdf(FILLED_UP_FORMS, AUTH_LETTER_OUTPUT_FILE_NAME + guestName + "_"+ confirmationCode + ".pdf")
    healthDecPDF = ProcessPdf(FILLED_UP_FORMS, HEALTH_DECLARATION_OUTPUT_FILE_NAME + guestName + "_"+ confirmationCode + ".pdf")
    # gatePassCleaningPDF = ProcessPdf(FILLED_UP_FORMS, GATEPASS_CLEANING_OUTPUT_FILE_NAME + "_"+ confirmationCode + ".pdf")
    # workPermitCleaningPDF = ProcessPdf(FILLED_UP_FORMS, WORK_PERMIT_CLEANING_OUTPUT_FILE_NAME + "_"+ confirmationCode + ".pdf")
    
    airBNBCheckinDate = pdf.massageCSVData(guest[csvHeader[1]])
    airBNBCheckinDateFormatted = datetime.strptime(airBNBCheckinDate, '%d/%m/%Y')
    checkinDate = airBNBCheckinDateFormatted.strftime("%b") + " " + airBNBCheckinDateFormatted.strftime("%d") + ", " + airBNBCheckinDateFormatted.strftime("%Y")
    airBNBCheckoutDate = pdf.massageCSVData(guest[csvHeader[2]])
    airBNBCheckoutDateFormatted = datetime.strptime(airBNBCheckoutDate, '%d/%m/%Y')
    checkoutDate = airBNBCheckoutDateFormatted.strftime("%b") + " " + airBNBCheckoutDateFormatted.strftime("%d") + ", " + airBNBCheckinDateFormatted.strftime("%Y")

    deltaBeforeDays = timedelta(7)
    issueDateWithDelta = airBNBCheckinDateFormatted - deltaBeforeDays

    data = {
        'guest1': pdf.massageCSVData(guest[csvHeader[0]]),
        'checkin': checkinDate,
        'checkout': checkoutDate,
        'issueDay': pdf.date_extention(int (issueDateWithDelta.strftime("%d"))),
        'issueMonth': issueDateWithDelta.strftime("%B"),
        'issueYear': issueDateWithDelta.strftime("%y"),
        'workPermitRequestDate': issueDateWithDelta.strftime("%B") + " "+ issueDateWithDelta.strftime("%d") + ", " + issueDateWithDelta.strftime("%y"),
        'workPermitDate': checkinDate

    }
    pdf.add_data_to_pdf(TEMPLATE_PATH + AUTH_LETTER_TEMPLATE_FILENAME, data)
    healthDecPDF.add_data_to_pdf(TEMPLATE_PATH + HEALTH_DECLARATION_TEMPLATE_FILENAME, data)
    #gatePassCleaningPDF.add_data_to_pdf(TEMPLATE_PATH + GATEPASS_CLEANING_TEMPLATE_FILENAME, data)
    #workPermitCleaningPDF.add_data_to_pdf(TEMPLATE_PATH + WORK_PERMIT_CLEANING_TEMPLATE_FILENAME, data)