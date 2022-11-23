"""
Will install pypiwin32 and pypandoc if not already installed.

"""

header_length = 11  # Number of items in the CSV header
table_cols = ('CRN', 'Subject and Course Number', 'Student ID', 'Student Last Name', 'Student First Name')

import csv
import os
import pathlib

try:
    import win32com.client as win32
except ModuleNotFoundError:
    os.system('python -m pip install pypiwin32')

try:
    import pypandoc
except ModuleNotFoundError:
    os.system('python -m pip install pypandoc')

# UI variables
LIMIT = 9001

# Test values
pwd = pathlib.Path(__file__).parent
subject = 'Enter email subject line'
email_file = pwd / 'test_template.txt'
css_file = pwd / 'style.css'
csv_file = pwd / 'test_csv.csv'
attachment_file = pwd / 'test_attachment.pdf'

def main(subject, email_file, css_file, csv_file, attachment_file, test):
    # Make column name strings globally available
    global sid, sfirst, smid, slast, flast, ffirst, femail, crn, title, course, term
    if test: test_message()

    # Read CSV
    cols, students_by_femail = read_csv(csv_file)
    sid, sfirst, smid, slast, flast, ffirst, femail, crn, title, course, term = cols[:header_length]

    # Make emails
    style = read_style(css_file)
    print(header := f'student count    recipient')
    print('`'*(20 + len(header)))
    student_count = 0
    for send_count, (femail, students) in enumerate(students_by_femail.items()):
        body = format_email_body(email_file, style, students)
        if not test:
            send_email(femail, subject, body, attachment_file)
        student_count += len(students)
        print(f'{len(students):^13}    {femail}')
    print()
    print(f'{send_count+1:>3} emails sent')
    print(f'{student_count:>3} students')

def read_csv(csv_file):
    with open(csv_file) as file:
        rows = csv.reader(file)
        cols = next(rows)
        email_col = cols[6]  # email column string
        return cols, make_students_by_professor(email_col, cols, rows)

def read_style(css_file):
    with open(css_file) as file:
        return ''.join(file.readlines())

def format_email_body(email_file, style, students):
    with open(email_file) as file:
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

def send_email(to, subject, body, attachment):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.HTMLBody = body
    mail.Attachments.Add(attachment)
    mail.Send()

def test_message():
    print(30*'-')
    print(f'testing from {__name__}')
    print(30*'-')

if __name__ == '__main__':
    main(subject, email_file, css_file, csv_file, attachment_file, test=True)

