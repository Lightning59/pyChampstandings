""" main file - Run this should prompt user for a properly formatted Excel
file then update that file with pretty output of championship standings"""

import tkinter
import tkinter.filedialog
import openpyxl


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
        lraw = len(rawstr)
        lsummary = len(summarystr)
        for sheet in wb.sheetnames:
            if sheet[:lraw]==rawstr:
                self.rawdatasheets.append(sheet)
            elif sheet[:lsummary]==summarystr:
                self.summarysheets.append(sheet)
            else:
                self.othersheets.append(sheet)

        # Wipe old summary sheets
        for sheet in self.summarysheets:
            self.wb.remove(self.wb[sheet])
        self.summarysheets=[]


    def init_outsheets(self):
        for sheet in self.rawdatasheets:
            if sheet not in self.summarysheets:
                self.wb.create_sheet(sheet.replace(self.rawstr,self.summarystr))
                self.summarysheets.append(sheet)


    def writeout(self):
        self.wb.save('out.xlsx')




if __name__ == '__main__':
    RAWNAMEBASE = "rawdata_"
    SUMMARYNAMEBAE = "summary_"

    the_wb = champWB(get_wb(),RAWNAMEBASE, SUMMARYNAMEBAE)
    the_wb.init_outsheets()
    the_wb.writeout()