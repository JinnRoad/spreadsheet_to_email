#	Spreadsheet mail merge project for SOAR

##	Directory contents

Documentation

-	README.md   # This file
-	todo.md     # A list of to-do's for the project


Program files

-	main.py     # Entry point for script
-	gui.py      # Runs GUI and collects document names
-	emailer.py  # Emailer tranforms the documents into emails, which are then sent
-	style.css   # Default CSS file

Test files for running the code

-	test_attachment.pdf
-	test_csv.csv
-	test_template.txt


##	Project plan

1.	Begin with
	1.	CSV containing student and teacher data (extracted via from a school database).
	1.	Email template (markdown)
2.	Convert to a mass email where each email is formatted by programmatically slicing the CSV into a table depending on the email recipient.
	-	Consider an excel pivot table of student records organized by professor and class.
	-	Each email has subsection of that pivot table.

The above actions were previously performed using the Microsoft Word "Mail Merge" function, however the tables had to manually be created by staff.
This process had to be completed more than once a semester, each time requiring fifty hours of tedious labor.

###	Pseudocode

	Level 0 Outline
		Collect data files
		Convert data files into HTML files containing a formatted email
		Send each HTML file as an email via Microsoft Outlook


	Level 1 Outline
		Open GUI to collect information
			Subject line
			Email template file
			Email template CSS file
			CSV file
			Attachment file
		Open clicking SEND, begin processing the files to send emails
			Read the CSV file and organize it into a dictionary by email recipient (faculty email)
			Read the CSS file
			Use the pandoc to format the email into HTML, guided by the CSS file
				Formatting includes inserting the faculty name and the table of students organized by class.
			Send the HTML file by hooking into Outlook via the win32com library.

###	Modules used

-	csv: Reading CSVs.
-	os: Installing libraries for the user if they aren't already installed.
-	pathlib: Sensible filepath syntax.
-	pypandoc: Conversion from Markdown to HTML.
-	tkinter: Creating GUI.
-	win32: Hooking into Microsoft Outlook to send emails.

