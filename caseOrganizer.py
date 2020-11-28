#! /usr/bin/env python3
# caseOrganizer.py - my father operates at many hospitals and needs me to
# collect all the cases from a specific hospital from 2013 - 2019 with the
# record of each year being in sepeate word documents.
# The collected cases need to be neatly added to a new word document
# labeling each year

# The format of the word document is as follows where each line is one patient:
# 1. Smith, B 111111 AA 1/1 "resp failutre" "chest tube removal"

import re, docx

# Case regex -> could not get comments to work in regex so here it is typed neatly
# caseRegex = re.compile(r"""(
#    ([a-zA-z-]*?[ ]??[a-zA-Z-]+,\s?[a-zA-Z]*) # name
#    \s*                                       # misc space
#    (\d{5,7})                                 # patient id number
#    \s*                                       # misc space
#    ([a-zA-Z]+)                               # hospital abbreviation
#    \s*                                       # misc space
#    (\d+/\d+)                                 # date
#    \s*                                       # misc space
#    ([a-zA-Z0-9,#/?‘’ +-–()–]+)               # problem
#    \s*                                       # misc space
#    ([a-zA-Z0-9,#\/?‘’ +-–()–]+)              # operation
#)""", re.VERBOSE)

caseRegex20132014 = re.compile(r'''([a-zA-z-]*?[ ]??[a-zA-Z-]+,\s?[a-zA-Z]*)\s*(\d{5,7})[ \t]*([a-zA-Z]+)\s*(\d+/\d+)\s*([a-zA-Z0-9,#/?‘’ +-/(/)–]+)\s*([a-zA-Z0-9,#/?‘’ +-/(/)–]+)''')
caseRegex20152019 = re.compile(r'''([a-zA-z-]*?[ ]??[a-zA-Z-.]+,\s?[a-zA-Z]*)\s*(\d{5,7})\s*(\d+/\d+)\s*([a-zA-Z]+)\s*([a-zA-Z0-9,#/?‘’ +-/(/)–]+)\s*([a-zA-Z0-9,#/?‘’ +-/(/)–]+)''')

# Data structure is year maps to case number which maps to all the other data
# other data: (Name, Patient #, Date, Hospital, Problem, Operation)
cases = {}

# Word documents called '20## CASES.docx'
# Reading word douments
for year in range (2013, 2020):
    print('Reading ' + str(year) + ' CASES.docx...')
    cases.setdefault(year, {})
    doc = docx.Document('/Users/Gabriel/Documents/CaseOrganization/'+ str(year) + ' CASES.docx')
    # Reading each line of the word document
    countNumberOfCases = 1
    for lines in range (0, len(doc.paragraphs)):
        if year == 2013 or year == 2014:
            casesGrouped = caseRegex20132014.search(doc.paragraphs[lines].text)
        else:
            casesGrouped = caseRegex20152019.search(doc.paragraphs[lines].text)
        # Hospital is Georgetown University or GU in the word document
        if casesGrouped != None and (casesGrouped.group(3) == 'GU' or casesGrouped.group(4) == 'GU'):
            cases[year].setdefault(countNumberOfCases, {})
            if year == 2013 or year == 2014:
                cases[year][countNumberOfCases] = {'Name': casesGrouped.group(1),
                                                  'ID': casesGrouped.group(2),
                                                  'Hospital': casesGrouped.group(3),
                                                  'Date': casesGrouped.group(4),
                                                  'Problem': casesGrouped.group(5),
                                                  'Operation': casesGrouped.group(6)}
            else:
                cases[year][countNumberOfCases] = {'Name': casesGrouped.group(1),
                                                  'ID': casesGrouped.group(2),
                                                  'Hospital': casesGrouped.group(4),
                                                  'Date': casesGrouped.group(3),
                                                  'Problem': casesGrouped.group(5),
                                                  'Operation': casesGrouped.group(6)}
            countNumberOfCases += 1

# Writing word document
print('Writing result docx called GU CASES...')
doc = docx.Document()
for year in range (2013, 2020):
    doc.add_paragraph(str(year)+ ' Cases\n')
    for caseNumber in range (1, len(cases[year]) + 1):
        paraObj = doc.add_paragraph(str(caseNumber) + '. ')
        for key in cases[year][caseNumber]:
            paraObj.add_run(str(key) + ': ' + cases[year][caseNumber][key] + '\t')
doc.save('/Users/Gabriel/Documents/CaseOrganization/GU CASES.docx')
print('DONE')
