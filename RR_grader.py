# import the packages that let me find files and export to xlsx
import os, pandas as pd, openpyxl

# find all downloaded Reading Response canvas pages
files = [file for file in os.listdir() if file.endswith('.html')]

# create a set of student names
students = set()

# filter non-names
def is_name(name):
    number  = all(chr == 'I' for chr in name)
    pronoun = '/' in name
    suffix  = name == 'Jr.' or name == 'Sr.'
    return not(number or pronoun or suffix)

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
            names = line.split()

            # the first and last words that are names will be used
            names = list(filter(is_name,names))
            students.add(names[0] + ' ' + names[-1])

# alphabetize by last name
def lastname(name):
    return name.split()[-1] # last word in name

students = list(students)
students.sort(key=lastname) # sort using the function lastname (default sort for strings is alphabetical)


# a datatype for discussion with a types, aliases, a list of dates, and student responses for them.
class Discussion:
    def __init__(self, abbr, aliases, dates, response):
        self.abbr     = abbr
        self.aliases  = aliases
        self.dates    = dates
        self.response = response
        self.db       = None

    def matches_alias(self, filename):
        return any((alias in filename) for alias in self.aliases)

discs = []
with open('categories.txt', 'r') as file:
    for line in file:

        # letters before the colon are used as the discussion type abbreviation
        abbr, aliases = line.split(':')

        # the abbreviation & any words after the colon are interpreted as an acceptable alias
        aliases = [abbr] + aliases.split()

        # create a mapping of student names to a list of whether or not they did each reading response
        response = {}
        for student in students:
            response[student] = []

        # create a list for the dates of discussions of this type
        dates = []

        # create a discussion of this new type
        discs.append(Discussion(abbr, aliases, dates, response))


# be able to pick out the date from the filename assuming it appears in the format '... Sep 12 ...'
months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
def Date(file):
    month = next(month for month in months if month in file)
    i = file.index(month)
    # returns the first three letters of the month and any numbers in the 8 characters following it
    file = file[i:i+8]
    return file[:3] + ' ' +  ''.join(c for c in file if c.isdigit())


# for each reading
for file in files:

    # import the raw html as a readable text document...
    with open(file, 'r', encoding='utf8') as data:
        data = data.read()

        # first type of discussion that matches the aliases in the filename
        disc = next(disc for disc in discs if disc.matches_alias(file))

        disc.dates.append(Date(file))

        # and scan the text for each student's name appearing in an author tag on a discussion post.
        for student in students:
            disc.response[student].append('X' if 'title="Author&#39;s name">' + student.split()[0] in data and student.split()[-1] in data else '')

# sum 'em all up
for disc in discs:
    disc.dates.append('Total')

    for student in students:
        # add up all the responses for the student that aren't blank 
        # ("if r" returns False if r is blank, otherwise True)
        disc.response[student].append(sum(1 for r in disc.response[student] if r))

# put them into a spreadsheet-like format
for disc in discs:
    disc.db = pd.DataFrame(disc.response.values(), index=disc.response.keys(), columns=disc.dates)
    print(disc.db)

# save as xlsx file
with pd.ExcelWriter('discussion_posts.xlsx') as writer:
    for disc in discs:
        # make a new sheet within the xlsx file titled with all the aliases of the response type
        disc.db.to_excel(writer, sheet_name = ' '.join(disc.aliases))

# do some funky stuff to make the Name column wider
book = openpyxl.load_workbook('discussion_posts.xlsx')

for disc in discs:
    book[' '.join(disc.aliases)].column_dimensions['A'].width = 20

book.save('discussion_posts.xlsx')

