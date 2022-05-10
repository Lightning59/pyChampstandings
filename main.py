""" main file - Run this should prompt user for a properly formatted Excel
file then update that file with pretty output of championship standings"""

import tkinter
import tkinter.filedialog
import openpyxl
import openpyxl.worksheet
import openpyxl.utils
import openpyxl.styles

boldGold = openpyxl.styles.Font(bold=True,color="FFD700")
BoldSilver = openpyxl.styles.Font(bold=True,color="C0C0C0")
BoldBronze = openpyxl.styles.Font(bold=True,color="CD7F32")
BlodBlack = openpyxl.styles.Font(bold=True)
Droppedgrey = openpyxl.styles.Font(color="D3D3D3")

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

    def calc_graphicRange(self, sheet, numracers, numweeks):
        # 3 lines worth of headers
        output_height = 3 + numracers
        # 2X num weeks covers driver and pts column each week then need to add positions
        output_width = 2 * numweeks
        for i in range(numweeks):
            output_width += i + 1
        ws = self.wb[sheet]
        endletter = openpyxl.utils.cell.get_column_letter(output_width)
        endcell = endletter + str(output_height)
        graphic_range = ws['A1':endcell]
        return graphic_range

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

        self.write_basic_headers(sheet,numweeks,numracers,sheet.replace("summary_",''))

    def write_basic_headers(self, sheet: str, numweeks: int, numracers: int, seriesname: str):
        graphicRange=self.calc_graphicRange(sheet,numracers,numweeks)
        graphicRange[0][0].value=seriesname

        week00cells=[0]
        for i in range(1,numweeks):
            week00cells.append(week00cells[i-1]+2+i)

        GoldFill = openpyxl.styles.PatternFill(start_color='FFD700', end_color='FFD700',
                                                fill_type='solid')  # change to FFFFFF
        SilverFill = openpyxl.styles.PatternFill(start_color='C0C0C0', end_color='C0C0C0',
                                                fill_type='solid')  # change to FFFFFF
        BronzeFill = openpyxl.styles.PatternFill(start_color='CD7F32', end_color='CD7F32',
                                                fill_type='solid')  # change to FFFFFF
        week=1
        for cellindex in week00cells:
            graphicRange[1][cellindex].value="Week "+str(week)
            week+=1

            graphicRange[2][cellindex].value="Driver"
            graphicRange[2][cellindex+1].value="Pts"
            graphicRange[2][cellindex+2].value="Finishes"

            graphicRange[3][cellindex].fill=GoldFill
            graphicRange[4][cellindex].fill=SilverFill
            graphicRange[5][cellindex].fill=BronzeFill

    def writeout(self):
        self.wb.save('out.xlsx')

class Series(object):
    def __init__(self,sheet,dropweeks):
        self.sheet=sheet
        self.drivers=[]
        self.weeks=[]
        self.dropweeks=dropweeks
        self.weeklysorteddirvers=[]

    def readdrivers(self):
        cvalue="start"
        i=3
        while True:
            cvalue=self.sheet.cell(row=1,column=i).value
            if cvalue==None:
                break
            d=Driver(cvalue,i)
            assert len(self.weeks)!=0
            d.readresults(self.sheet,len(self.weeks))
            self.drivers.append(d)
            i+=1

    def readweeks(self):
        cvalue="start"
        i=2
        while True:
            cvalue=self.sheet.cell(row=i,column=1).value
            if cvalue==None:
                break
            self.weeks.append(cvalue)
            i+=1

    def sortdrivers(self,week):
        sordr=self.drivers.copy()
        # sort by number of points for the week. If that is not enought then sort based on best dropped position going through all dropped positions as needed. - Need to still add Alpha drivers name last?
        sordr=sorted(sordr,key=lambda y: [-y.weeklypoints[week]]+[y.weeklydropped[week][p][1] for p in range(len(y.weeklydropped))])
        # I belieive this will throw errors if it trys to compare 'DNP' with an int. Need to add another int in the tuple where anything other than a number becomes 0
        # or could make these poition object then define the __lt__ for position class?
        self.weeklysorteddirvers.apppend(sordr)


class Driver(object):
    def __init__(self,name,col):
        self.name=name
        self.col=col
        self.allresults=[]
        self.thisweekresults=[]
        self.weeklyresults=[]
        self.weeklypoints=[]
        self.weeklydropped=[]

    def readresults(self,sheet,numweeks):
        for i in range(2,numweeks+2):
            self.allresults.append((i-1,sheet.cell(row=i,column=self.col).value))

    def calcpoints_results_byweek(self,numweeks,dropweeks):
        weekres=self.allresults[0:numweeks]
        droppedres=[]

        for i in range(len(weekres)):
            if weekres[i][1] in POINTS.keys():
                weekres[i]=(weekres[i][0],weekres[i][1],POINTS[weekres[i][1]])
            else:
                weekres[i] = (weekres[i][0], weekres[i][1], 0)


        weekres = sorted(weekres,key=lambda y: y[2], reverse=True)

        if numweeks>dropweeks:
            for i in range(len(weekres)):
                if i<(numweeks-dropweeks):
                    weekres[i] = (weekres[i][0], weekres[i][1], weekres[i][2],1)
                else:
                    weekres[i] = (weekres[i][0], weekres[i][1], weekres[i][2], 0)
                    droppedres.append((weekres[i][0], weekres[i][1], weekres[i][2], 0))
        else:
            weekres[0] = (weekres[0][0], weekres[0][1], weekres[0][2], 1)
            for i in range(1,len(weekres)):
                weekres[i] = (weekres[i][0], weekres[i][1], weekres[i][2], 0)
                droppedres.append((weekres[i][0], weekres[i][1], weekres[i][2], 0))

        #sort all the dropped results in order of quality as we may need them to break ties.
        droppedres=sorted(droppedres,key=lambda y: (type(y[1])==str,y[1]))
        self.weeklyresults.append(weekres)
        self.weeklydropped.append(droppedres)
        score=0
        for item in weekres:
            if item[3]==1:
                score+=item[2]
        self.weeklypoints.append(score)






POINTS={1:25,2:18,3:15,4:12,5:10,6:8,7:6,8:4,9:2,10:1}

if __name__ == '__main__':
    RAWNAMEBASE = "rawdata_"
    SUMMARYNAMEBAE = "summary_"

    the_wb = champWB(get_wb(), RAWNAMEBASE, SUMMARYNAMEBAE)
    the_wb.init_outsheets()
    the_wb.format_outsheet("summary_Clio", 3, 5)
    x=Series(the_wb.wb['rawdata_Clio'],2)
    x.readweeks()
    x.readdrivers()
    for driver in x.drivers:
        driver.calcpoints_results_byweek(3,2)
    x.sortdrivers(0)
    the_wb.writeout()
