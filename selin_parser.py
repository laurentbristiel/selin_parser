from openpyxl import load_workbook
import sys

class selin_parser:
    """
    INSTALL
    Those steps are customized for Windows - but tested only on Mac OS X
    - install Python 2.x (last version)
        https://www.python.org/downloads/release/python-278/
    - install openpyxl library
        in a DOS : easy_install openpyxl
    - put this file (selin_parser.py) and the XLXS modifiers file in the same directory
    - in a DOS, launch:
        python selin_parser.py my_modifier_file.xlsx
    """

    def __init__(self):
        # row 4 contains the name of the modifiers
        self._modifiers_name_row = 4
        # this range contains the list of all the characters modifiers
        self.char_modifier_begin = 80
        self.char_modifier_end = 120
        # this range contains the list of all the other modifiers
        self.other_modifier_begin = 12
        self.other_modifier_end = 79
        # column 4 contains the name/code of the religions
        self._religion_name_col = 4
        # this range contains the list of all the religion
        self._religion_row_begin = 7
        self._religion_row_end = 279

    def parse_modifiers_sheet(self, file_path):

        wb = load_workbook(filename=file_path, data_only=True)
        ws = wb.get_sheet_by_name('Definition')

        for religion_row in range(self._religion_row_begin, self._religion_row_end):
            print "#################################################"
            print ws.cell(row=religion_row, column=self._religion_name_col).value
            print "#################################################"
            self.print_modifiers(ws, 'character_modifier', religion_row, self.char_modifier_begin, self.char_modifier_end)
            self.print_modifiers(ws, 'other_modifier', religion_row, self.other_modifier_begin, self.other_modifier_end)

    def print_modifiers(self, ws, range_name, religion_row, begin, end):
        line_prefix = '\t\t'
        print line_prefix + range_name + ' = {'
        for col in range(begin, end):
            header_value = ws.cell(row=self._modifiers_name_row, column=col).value
            cell_value = ws.cell(row=religion_row, column=col).value
            if header_value is None or cell_value == 0:
                continue
            if isinstance(cell_value, float):
                print line_prefix + '\t' + header_value.encode('utf-8') + " = %0.2f" % cell_value
            if isinstance(cell_value, int):
                print line_prefix + '\t' + header_value.encode('utf-8') + " = " + str(cell_value)
        print line_prefix + '}'

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print "usage: python selin_parser.py filename.xlsx"
    else:
        selin_parser = selin_parser()
        selin_parser.parse_modifiers_sheet(sys.argv[1])