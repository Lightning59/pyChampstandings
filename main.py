""" main file - Run this should prompt user for a properly formatted Excel
file then update that file with pretty output of championship standings"""

import tkinter
import tkinter.filedialog
import openpyxl
import openpyxl.worksheet
import openpyxl.utils
import openpyxl.styles


def get_wb():
    """Launches a file selection window. Returns a filepath string or raises a FileNotFoundError if user cancels"""
    rootwindow = tkinter.Tk()
    rootwindow.withdraw()

    file_path = tkinter.filedialog.askopenfilename()
    try:
        assert file_path != ''
    except AssertionError:
        print("You did not select a valid file.")
        raise FileNotFoundError
    try:
        wb = openpyxl.load_workbook(file_path)
    except:
        print("error opening workbook")
        raise NameError
    return wb


class champWB(object):

    def __init__(self, wb: openpyxl.Workbook, rawstr: str, summarystr: str):
        self.wb = wb
        self.rawdatasheets = []
        self.rawstr = rawstr
        self.summarysheets = []
        self.summarystr = summarystr
        self.othersheets = []
        self.series = []
        lraw = len(rawstr)
        lsummary = len(summarystr)
        for sheet in wb.sheetnames:
            if sheet[:lraw] == rawstr:
                self.rawdatasheets.append(sheet)
                self.series.append(sheet[lraw:])
            elif sheet[:lsummary] == summarystr:
                self.summarysheets.append(sheet)
            else:
                self.othersheets.append(sheet)

        # Wipe old summary sheets
        for sheet in self.summarysheets:
            self.wb.remove(self.wb[sheet])
        self.summarysheets = []

    def init_outsheets(self):
        '''create new blank summary sheets for each series.'''
        for sheet in self.rawdatasheets:
            if sheet not in self.summarysheets:
                self.wb.create_sheet(sheet.replace(self.rawstr, self.summarystr))
                self.summarysheets.append(sheet)

    def format_outsheet(self, sheet: str, numweeks: int, numracers: int):

        # 3 lines worth of headers
        output_height = 3 + numracers
        # 2X num weeks covers driver and pts column each week then need to add positions
        output_width = 2 * numweeks
        for i in range(numweeks):
            output_width += i + 1
        ws = self.wb[sheet]
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=output_width)

        startweekcol = 1
        for i in range(1, numweeks + 1):
            ws.merge_cells(start_row=2, start_column=startweekcol, end_row=2, end_column=startweekcol + i + 1)
            if i != 1:
                ws.merge_cells(start_row=3, start_column=startweekcol + 2, end_row=3, end_column=startweekcol + 1 + i)
            startweekcol += i + 2

        endletter = openpyxl.utils.cell.get_column_letter(output_width)
        endcell = endletter + str(output_height)
        graphic_range = ws['A1':endcell]
        self.graphic_range = graphic_range

        # center and white background fill all cells
        whiteFill = openpyxl.styles.PatternFill(start_color='00FFFF', end_color='00FFFF',
                                                fill_type='solid')  # change to FFFFFF
        centerall = openpyxl.styles.alignment.Alignment(horizontal="center", vertical="center")
        for row in graphic_range:
            for cell in row:
                cell.alignment = centerall
                cell.fill = whiteFill

        thickblack = openpyxl.styles.Side(border_style='medium', color="000000")
        thinblack = openpyxl.styles.Side(border_style='thin', color="000000")
        thick_allsides_border = openpyxl.styles.Border(left=thickblack, right=thickblack, top=thickblack,
                                                       bottom=thickblack)
        # First Two Headers for series and weeks - Set Borders

        bold14pt = openpyxl.styles.Font(size=14, bold=True)
        boldonly = openpyxl.styles.Font(bold=True)
        for cell in graphic_range[0]:
            cell.border = thick_allsides_border
            cell.font = bold14pt
        for cell in graphic_range[1]:
            cell.font = boldonly
            cell.border = thick_allsides_border

        # Cacluate indicies for week breaks
        i_week_breaks = []
        prevstop = 0
        for i in range(numweeks):
            i_week_breaks.append(prevstop + 2 + i + 1)
            prevstop += 3 + i

        # Format row 3 with thick upper lines and thick start stop week breaks. Thin lower line
        row3 = graphic_range[2]
        for i in range(len(row3)):
            if i == 0:
                row3[i].border = openpyxl.styles.Border(left=thickblack, top=thickblack, bottom=thinblack)
            elif i == output_width - 1:
                row3[i].border = openpyxl.styles.Border(right=thickblack, top=thickblack, bottom=thinblack)
            elif i in i_week_breaks:
                row3[i].border = openpyxl.styles.Border(left=thickblack, top=thickblack, bottom=thinblack)
            elif i + 1 in i_week_breaks:
                row3[i].border = openpyxl.styles.Border(right=thickblack, top=thickblack, bottom=thinblack)
            else:
                row3[i].border = openpyxl.styles.Border(top=thickblack, bottom=thinblack)

        # format las row ith Thick bottom lin and thin top line otherwise Thick week breaks.
        rowlast = graphic_range[output_height - 1]
        for i in range(len(rowlast)):
            if i == 0:
                rowlast[i].border = openpyxl.styles.Border(left=thickblack, top=thinblack, bottom=thickblack)
            elif i == output_width - 1:
                rowlast[i].border = openpyxl.styles.Border(right=thickblack, top=thinblack, bottom=thickblack)
            elif i in i_week_breaks:
                rowlast[i].border = openpyxl.styles.Border(left=thickblack, top=thinblack, bottom=thickblack)
            elif i + 1 in i_week_breaks:
                rowlast[i].border = openpyxl.styles.Border(right=thickblack, top=thinblack, bottom=thickblack)
            else:
                rowlast[i].border = openpyxl.styles.Border(top=thinblack, bottom=thickblack)

        # Format all interim lines all thin upper and lower. Thick start, stop, weekbreak.
        for j in range(3, output_height - 1):
            for i in range(output_width):
                if i == 0:
                    graphic_range[j][i].border = openpyxl.styles.Border(left=thickblack, top=thinblack,
                                                                        bottom=thinblack)
                elif i == output_width - 1:
                    graphic_range[j][i].border = openpyxl.styles.Border(right=thickblack, top=thinblack,
                                                                        bottom=thinblack)
                elif i in i_week_breaks:
                    graphic_range[j][i].border = openpyxl.styles.Border(left=thickblack, top=thinblack,
                                                                        bottom=thinblack)
                elif i + 1 in i_week_breaks:
                    graphic_range[j][i].border = openpyxl.styles.Border(right=thickblack, top=thinblack,
                                                                        bottom=thinblack)
                else:
                    graphic_range[j][i].border = openpyxl.styles.Border(top=thinblack, bottom=thinblack)

    def writeout(self):
        self.wb.save('out.xlsx')


if __name__ == '__main__':
    RAWNAMEBASE = "rawdata_"
    SUMMARYNAMEBAE = "summary_"

    the_wb = champWB(get_wb(), RAWNAMEBASE, SUMMARYNAMEBAE)
    the_wb.init_outsheets()
    the_wb.format_outsheet("summary_Clio", 3, 5)
    the_wb.writeout()
