## CPL Repair Form

# This Document takes the users input to print a CPL Repair Form onto a sticky label
# Here we will import the required packages used in this script

from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import time
#from openpyxl import Workbook, load_workbook
import csv
from csv import writer
import os
import sys

def cplRepairForm():

    # Creating a function to define the needed repair that will be printed on the sheet
    def repairFunc(): # Function to fill the repairNeeded variable
        repair = int(input("Select the Repair:\n1. Bad or Missing RFID Tag\n2. Book Repair\n3. BCDs / CDs / DVDs / Games\n4. Call / Location, Shelf Review, No Record Found\n5. Other Exceptions\nUser Selection: "))

        if repair == 1:
            badMissingRFIDTag = int(input("Select the Repair:\n1. Tag won't Read\n2. New Tag Needed\nUser Selection: "))
            if badMissingRFIDTag == 1:
                value = "Bad/Missing RFID Tag - Tag Won't Read"
                return value
            if badMissingRFIDTag == 2:
                value = "Bad or Missing RFID Tag - New Tag Needed"
                return value

        elif repair == 2:
            bookRepair = int(input("Select the Repair:\n1. Loose or Torn Page(s)\n2. Mylar Jacket / Cover\n3. New Spine Label (Alpha, Genre/Spine)\n4. Worn Item/Retire\n5. Damage Noted Sticker - Needed Approval\nUser Selection: "))
            if bookRepair == 1:
                pageNumber = input("What page(s) are damaged?\n")
                value = "Loose or Torn Page(s) # "+ pageNumber
                return value
            elif bookRepair == 2:
                value = "Mylar Jacket / Cover"
                return value
            elif bookRepair == 3:
                labelRepair = int(input("What Spine Label Repair is needed?\n1. Alpha\n2. Genre / Spine\nUser Selection: "))
                if labelRepair == 1:
                    value = "Label - Alpha"
                    return value
                elif labelRepair == 2:
                    value = "Label - Genre / Spine"
                    return value
            
            elif bookRepair == 4:
                value = "Worn Item/ Retire"
                return value
        
            elif bookRepair == 5:
                value = "Damage Noted Sticker - Needed / Approval"
                return value

        elif repair == 3:
            discRepair = int(input("What Disc Repair is needed?\n1. Will Not Play / Scratched\n2. Replace Case\n3. Replace Artwork\nUser Selection: "))
            if discRepair == 1:
                value = "Will Not Play / Scratched"
                return value
        
            elif discRepair == 2:
                value = "Replace Case"
                return value
        
            elif discRepair == 3:
                value = "Replace Artwork"
                return value

        elif repair == 4:
            callLoc = int(input("What repair is needed?\n1. Call Location\n2. 5 Items or More on Shelf\n3. No Record Found\nUser Selection: "))
            if callLoc == 1:
                existing = input("Existing Call # or Location:\n")
                suggested = input("Suggested Call # or Location:\n")
                value = "Call # or Location\nExisting Call # or Location: " + existing + "\nSuggested Call # or Location: " + suggested
                return value
        
            elif callLoc == 2:
                value = "Shelf Review - 5 Items or More on Shelf"
                return value
        
            elif callLoc == 3:
                value = "No Record Found"
                return value

    
        elif repair == 5:
            value = "Other Exceptions"
            return value

    
    

    directoryName = os.path.dirname(__file__)
    template = os.path.join(directoryName, 'cplRepairFormTemplate.docx')
    document = MailMerge(template)

    ## Acquiring the needed data to fill the merge fields from the user

    staffInitials = input("Staff Initials: ")
    staffInitials = staffInitials.upper()
    todaysDate = '{:%b-%d-%Y}'.format(date.today())
    neededRepair = repairFunc()                  # function we created in line 17
    additionalNotes = input("Additional Notes: ")

    data_to_append = [staffInitials, todaysDate, neededRepair, additionalNotes]

    # Opening Excel Workbook
    #wb = load_workbook(r'C:\Users\brydenc\Documents\CPL Repair Form Revised\cplRepairData.xlsx')
    # Calling the active worksheet in the Excel Workbook
    #ws = wb.active

    # Here we are appending the user input data to the Excel Workbook and saving it
    #ws.append(data_to_append)
    #wb.save(r'C:\Users\brydenc\Documents\CPL Repair Form Revised\cplRepairData.xlsx')
    #wb.close()

        ## Putting the users information into the merge field
    document.merge(
        staffInitials = staffInitials.upper(),
        todaysDate = todaysDate,
        neededRepair = neededRepair,
        additionalNotes = additionalNotes
    )
 

    os.remove('cplRepairFormStickyLabel.docx')          # Clears file location in folder for new print job
    document.write('cplRepairFormStickyLabel.docx')     # Saves sticky label to be printed in folder
    #os.startfile('cplRepairFormStickyLabel.docx', 'print')    # Opens sticky label and prints in created Word document
    os.startfile('cplRepairFormStickyLabel.docx')       # Just opens sticky label document and does not print, user must select print in Word

    # Appending new data to a csv file
    with open(r'C:\Users\Christopher\Documents\CPL Repair Form Revised\cplRepairData.csv', 'a') as csvfile:

        writer_object = writer(csvfile, lineterminator = '\n')
        writer_object.writerow(data_to_append)
        csvfile.close()


def runItBack():
    restart = str(input('Do you have another damaged item needing repair? (Y/N)\nUser Selection: '))
    restart = restart.upper()
    if restart == 'Y':
        cplRepairForm()
    else:
        print('CPL Repair Forms Script is now closing.')
        time.sleep(3)
        sys.exit()



cplRepairForm()
50 * runItBack() # Re-runs the function at most 50 times before it automatically terminates
