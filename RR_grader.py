# import the packages that let me find files and export to xlsx
import os, pandas as pd, openpyxl

# find all downloaded Reading Response canvas pages
files = [file for file in os.listdir() if file.endswith('.html')]

# create a set of student names
students = set()

# open the list of student names (txt file copied and pasted from the People tab of canvas)
with open('students.txt', 'r') as file:
    for line in file:

        # if the line doesn't contain any of the gunk that occurs when you try to copy the list directly into a txt file
        if (    'Name'      not in line
            and 'HTS'       not in line 
            and 'Student'   not in line 
            and 'Teacher'   not in line 
            and 'Stoneman'  not in line
            and ':'         not in line
            and not any(chr.isdigit() for chr in line) ):

            # split the line into a list of words separated by spaces
            name = line.split()

            # the first and last word in the name line will be used (assuming the last one is not a pronoun '(He/Him)' or a suffix 'I')
            students.add(name[0] + ' ' + (name[-1] if len(name[-1]) > 1 and '/' not in name[-1] else name[-2]))

# alphabetize by last name
students = list(students)
students.sort(key=lambda name: name.split()[-1])

# create a mapping of student names to a list of whether or not they did each reading response
RR = {}
ACD = {}
for student in students:
    RR[student] = []
    ACD[student] = []

# for each reading
for file in files:

    # import the raw html as a readable text document...
    with open(file, 'r', encoding='utf8') as data:
        data = data.read()

        # pick type of discussion post based on file title
        response = RR if 'Reading' in file else ACD

        # and scan the text for each student's name appearing in an author tag on a discussion post.
        for student in students:
            response[student].append('X' if 'title="Author&#39;s name">' + student.split()[0] in data and student.split()[-1] in data else '')

# be able to pick out the date from the filename assuming it appears in the format '... Sep 12 ...'
months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
def Date(file):
    month = next(month for month in months if month in file)
    i = file.index(month)
    # returns the first three letters of the month and any numbers in the 8 characters following it
    file = file[i:i+8]
    return file[:3] + ' ' +  ''.join(c for c in file if c.isdigit())

# abbreviate the names of the readings
readings = ['RR ' + Date(file) for file in files if 'Reading' in file or 'RR' in file]
discussions = ['ACD ' + Date(file) for file in files if 'After' in file or 'ACD' in file]

# sum 'em all up
readings.append('Total')
discussions.append('Total')

for student in students:
    RR[student].append(sum(1 for r in RR[student] if r))
    ACD[student].append(sum(1 for r in ACD[student] if r))

# put them into a spreadsheet-like format
rr = pd.DataFrame(RR.values(), index=RR.keys(), columns=readings)
acd = pd.DataFrame(ACD.values(), index=ACD.keys(), columns=discussions)

print(rr)
print(acd)

# shorten the column names for Excel
rr.columns = [read[3:] for read in readings[:-1]] + ['Total']
acd.columns = [disc[4:] for disc in discussions[:-1]] + ['Total']

# save as xlsx file
with pd.ExcelWriter('discussion_posts.xlsx') as writer:
    rr.to_excel(writer, sheet_name='Reading Responses')
    acd.to_excel(writer, sheet_name='After Class Discussions')

# do some funky stuff to make the Name column wider
book = openpyxl.load_workbook('discussion_posts.xlsx')

book['Reading Responses'].column_dimensions['A'].width = 20
book['After Class Discussions'].column_dimensions['A'].width = 20

book.save('discussion_posts.xlsx')

