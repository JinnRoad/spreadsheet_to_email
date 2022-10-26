"""
COMMAND LINE USE

COMMAND FOR TESTING (TODO needs revision)
python emailer.py ~/Documents/template.txt ~/Documents/test.csv ~/Documents/attachment.pdf

# Requires pandoc installed to convert markdown to html
# Will install pypiwin32 automatically

    python emailer.py template.txt file.csv attachment_file.xxx [TEST_RECIPIENT@TEST.COM]

For the test recipient, I made a personl outlook account.

"""

# TODO Put template into separate file
# TODO Add readme and make usable by others
# TODO Build GUI for tkinter | kivy
# TODO Revise instructions once GUI is built
# TODO Get subject line from GUI

# USER INTERFACE


header_length = 11  # Number of items in the CSV header
table_cols = ('CRN', 'Subject and Course Number', 'Student ID', 'Student Last Name', 'Student First Name')



# CODE

import csv
import os
import sys

try:
    import win32com.client as win32
except ModuleNotFoundError:
    os.system('python -m pip install pypiwin32')

try:
    import pypandoc
except ModuleNotFoundError:
    os.system('python -m pip install pypandoc')

# UI variables
#template, csv, attachment = sys.argv[1:4]
TESTING = True
LIMIT = 9001
template_file = 'test_template.txt'
style_file = 'style.css'
csv_file = 'test_csv.csv'
attachment = 'test_attachment.pdf'

def main():
    # Make column name strings globally available
    global sid, sfirst, smid, slast, flast, ffirst, femail, crn, title, course, term

    # Read CSV
    cols, students_by_femail = read_csv(csv_file)
    sid, sfirst, smid, slast, flast, ffirst, femail, crn, title, course, term = cols[:header_length]

    # Make emails
    style = read_style(style_file)
    for send_count, (femail, students) in enumerate(students_by_femail.items()):
        if send_count == LIMIT:
            break
        email = format_template(template_file, style, students)
        #email = TEST_EMAIL if TESTING else femail
        #send_email(email, subject, body)
        print(f'recipient: {femail:50}student count: {len(students)}')
    print(f'\ntotal emails sent: {send_count+1}')

def read_csv(csv_file):
    with open(csv_file) as file:
        rows = csv.reader(file)
        cols = next(rows)
        email_col = cols[6]  # email column string
        return cols, make_students_by_professor(email_col, cols, rows)

def read_style(style_file):
    with open(style_file) as file:
        return ''.join(file.readlines())

def format_template(template_file, style, students):
    with open(template_file) as file:
        template = ''.join(file.readlines())
    format_dict = students[0]
    format_dict['table'] = make_table(students)
    email = template.format(**format_dict)
    return style + md2html(email)

def make_students_by_professor(email_col, cols, rows):
    students_by_femail = {}
    for row in rows:
        student = dict(zip(cols, row))
        femail = student[email_col]
        if femail not in students_by_femail:
            students_by_femail[femail] = [student]
        else:
            students_by_femail[femail].append(student)

    # Sort the student list by CRN
    for femail in students_by_femail:
        students_by_femail[femail] = sorted(students_by_femail[femail], key=lambda x: x['CRN'])

    return students_by_femail

def make_body(femail, students):
    greeting_ = greeting.format(**students[0])
    table = make_table(students)
    return style + md2html('\n\n'.join((greeting_, intro, table, outro)))

def md2html(string):
    return pypandoc.convert_text(string, 'html', format='md')

def make_table_(students):
    header = make_row(dict(zip(table_cols, table_cols)))
    rows = '\n'.join([header] + [make_row(student) for student in students])
    return f'<table>\n{rows}\n</table>'

def make_row_(student):
    values = (student[col] for col in table_cols)
    return '<tr>' + ''.join(f'<td>{value}</td>' for value in values) + '</tr>'

def make_table(students):
    header = make_row(dict(zip(table_cols, table_cols)))
    hline = make_row(dict(zip(table_cols, ['-']*len(table_cols))))
    rows = [make_row(student) for student in students]
    return '\n'.join([header] + [hline] + rows)

def make_row(student):
    return '|'.join(student[col] for col in table_cols)

def send_email(to, subject, body):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.HTMLBody = body
    mail.Attachments.Add(attachment)
    mail.Send()

main()
