""" main file - Run this should prompt user for a properly formatted Excel
file then update that file with pretty output of championship standings"""

import tkinter
import tkinter.filedialog
import openpyxl
import openpyxl.worksheet
import openpyxl.utils


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

    def __init__(self,wb: openpyxl.Workbook,rawstr:str, summarystr:str):
        self.wb = wb
        self.rawdatasheets = []
        self.rawstr=rawstr
        self.summarysheets = []
        self.summarystr=summarystr
        self.othersheets = []
        self.series=[]
        lraw = len(rawstr)
        lsummary = len(summarystr)
        for sheet in wb.sheetnames:
            if sheet[:lraw]==rawstr:
                self.rawdatasheets.append(sheet)
                self.series.append(sheet[lraw:])
            elif sheet[:lsummary]==summarystr:
                self.summarysheets.append(sheet)
            else:
                self.othersheets.append(sheet)

        # Wipe old summary sheets
        for sheet in self.summarysheets:
            self.wb.remove(self.wb[sheet])
        self.summarysheets=[]


    def init_outsheets(self):
        '''create new blank summary sheets for each series.'''
        for sheet in self.rawdatasheets:
            if sheet not in self.summarysheets:
                self.wb.create_sheet(sheet.replace(self.rawstr,self.summarystr))
                self.summarysheets.append(sheet)

    def format_outsheet(self,sheet:str,numweeks:int,numracers:int):

        #3 lines worth of headers
        output_height=3+numracers
        #2X num weeks covers driver and pts column each week then need to add positions
        output_width=2*numweeks
        for i in range(numweeks):
            output_width+=i+1
        ws=self.wb.get_sheet_by_name(sheet)
        endletter=openpyxl.utils.cell.get_column_letter(output_width)
        endcell=endletter+str(output_height)
        graphic_range=ws['A1':endcell]


        print('test')







    def writeout(self):
        self.wb.save('out.xlsx')




if __name__ == '__main__':
    RAWNAMEBASE = "rawdata_"
    SUMMARYNAMEBAE = "summary_"

    the_wb = champWB(get_wb(),RAWNAMEBASE, SUMMARYNAMEBAE)
    the_wb.init_outsheets()
    the_wb.format_outsheet("rawdata_Clio",2,5)
    the_wb.writeout()