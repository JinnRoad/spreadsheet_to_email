# TODO Test sending out actual emails to your gmail
# TODO Build GUI with kivy
# TODO Somehow either make the header_length variable unnecessary or automatic
# TODO Is the limit variable necessary
# TODO Add readme and make usable by others

import gui
import emailer

def main():
    gui.main()
    emailer.main(
        gui.files['subject'],
        gui.files['email'],
        gui.files['css'],
        gui.files['csv'],
        gui.files['attachment'],
        gui.files['test'],
        )

main()
