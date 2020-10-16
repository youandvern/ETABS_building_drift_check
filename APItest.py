import os
import sys
import comtypes.client
import pandas as pd
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTextEdit, QFileDialog, QApplication, QLineEdit, QTableView
from PyQt5.QtCore import QAbstractTableModel, Qt



class pandasModel(QAbstractTableModel):
    # create pandas model class to display dataframe in QTableView
    # https://learndataanalysis.org/display-pandas-dataframe-with-pyqt5-qtableview-widget/

    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid() and role == Qt.DisplayRole:
            return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None


def is_number(s):
    # funtion to test if a string input is a number
    # https://stackoverflow.com/questions/354038/how-do-i-check-if-a-string-is-a-number-float
    try:
        float(s)
        return True
    except ValueError:
        return False

class EtabsModel:
    # my ETABS API class to open and manipulate etabs model
    def __init__(self, modelpath, etabspath="C:/Program Files/Computers and Structures/ETABS 17/ETABS.exe", existinstance=False, specprogpath=False):
        # set the following flag to True to attach to an existing instance of the program
        # otherwise a new instance of the program will be started
        self.AttachToInstance = existinstance

        # set the following flag to True to manually specify the path to ETABS.exe
        # this allows for a connection to a version of ETABS other than the latest installation
        # otherwise the latest installed version of ETABS will be launched
        self.SpecifyPath = specprogpath

        # if the above flag is set to True, specify the path to ETABS below
        self.ProgramPath = etabspath

        # full path to the model
        # set it to the desired path of your model
        self.FullPath = modelpath
        [self.modelPath, self.modelName] = os.path.split(self.FullPath)
        if not os.path.exists(self.modelPath):
            try:
                os.makedirs(self.modelPath)
            except OSError:
                pass

        if self.AttachToInstance:
            # attach to a running instance of ETABS
            try:
                # get the active ETABS object
                self.myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
                self.success = True
            except (OSError, comtypes.COMError):
                print("No running instance of the program found or failed to attach.")
                self.success = False
                sys.exit(-1)

        else:
            # create API helper object
            self.helper = comtypes.client.CreateObject('ETABSv17.Helper')
            self.helper = self.helper.QueryInterface(comtypes.gen.ETABSv17.cHelper)
            if self.SpecifyPath:
                try:
                    # 'create an instance of the ETABS object from the specified path
                    self.myETABSObject = self.helper.CreateObject(self.ProgramPath)
                except (OSError, comtypes.COMError):
                    print("Cannot start a new instance of the program from " + self.ProgramPath)
                    sys.exit(-1)
            else:

                try:
                    # create an instance of the ETABS object from the latest installed ETABS
                    self.myETABSObject = self.helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
                except (OSError, comtypes.COMError):
                    print("Cannot start a new instance of the program.")
                    sys.exit(-1)

            # start ETABS application
            self.myETABSObject.ApplicationStart()

        # create SapModel object
        self.SapModel = self.myETABSObject.SapModel

        # initialize model
        self.SapModel.InitializeNewModel()

        # create new blank model
        # ret = self.SapModel.File.NewBlank()

        # open existing model
        ret = self.SapModel.File.OpenFile(self.FullPath)

        """
        # save model
        print(self.FullPath)
        ret = self.SapModel.File.Save(self.FullPath)
        print(ret)
        print("ETABS mod - model saved")
        """

        # run model (this will create the analysis model)
        ret = self.SapModel.Analyze.RunAnalysis()

        # get all load combination names
        self.NumberCombo = 0
        self.ComboNames = []
        [self.NumberCombo, self.ComboNames, ret] = self.SapModel.RespCombo.GetNameList(self.NumberCombo, self.ComboNames)

        # isolate drift combos by searching for "drift" in combo name
        self.DriftCombos = []
        for combo in self.ComboNames:
            lowerCombo = combo.lower()
            # skip combinations without drift in name
            if "drift" not in lowerCombo:
                continue
            self.DriftCombos.append(combo)

        self.StoryDrifts = []
        self.JointDisplacements = []
        pd.set_option("max_columns", 8)
        # pd.set_option("precision", 4)

    def story_drift_results(self, dlimit=0.01):
        # returns dataframe drift results for all drift load combinations
        self.StoryDrifts = []
        for dcombo in self.DriftCombos:
            # deselect all combos and cases, then only display results for combo passed
            ret = self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
            ret = self.SapModel.Results.Setup.SetComboSelectedForOutput(dcombo)

            # initialize drift results
            NumberResults = 0
            Stories = []
            LoadCases = []
            StepTypes = []
            StepNums = []
            Directions = []
            Drifts = []
            Labels = []
            Xs = []
            Ys = []
            Zs = []

            [NumberResults, Stories, LoadCases, StepTypes, StepNums, Directions, Drifts, Labels, Xs, Ys, Zs, ret] = \
                self.SapModel.Results.StoryDrifts(NumberResults, Stories, LoadCases, StepTypes, StepNums, Directions,
                                                  Drifts, Labels, Xs, Ys, Zs)
            # append all drift results to storydrifts list
            for i in range(NumberResults):
                self.StoryDrifts.append((Stories[i], LoadCases[i], Directions[i], Drifts[i], Drifts[i] / dlimit))

        # set up pandas data frame and sort by drift column
        labels = ['Story', 'Combo', 'Direction', 'Drift', 'DCR(Drift/Limit)']
        df = pd.DataFrame.from_records(self.StoryDrifts, columns=labels)
        dfSort = df.sort_values(by=['Drift'], ascending=False)
        dfSort.Drift = dfSort.Drift.round(4)
        dfSort['DCR(Drift/Limit)'] = dfSort['DCR(Drift/Limit)'].round(2)
        return dfSort

    def story_torsion_check(self):
        # returns dataframe of torsion results for drift combinations
        self.JointDisplacements = []
        for dcombo in self.DriftCombos:
            # deselect all combos and cases, then only display results for combo passed
            ret = self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
            ret = self.SapModel.Results.Setup.SetComboSelectedForOutput(dcombo)

            # initialize joint drift results
            NumberResults = 0
            Stories = []
            LoadCases = []
            Label  = ''
            Names = ''
            StepType = []
            StepNum = []
            # Directions = []
            DispX = []
            DispY = []
            DriftX = []
            DriftY = []

            [NumberResults, Stories, Label, Names, LoadCases, StepType, StepNum, DispX, DispY, DriftX, DriftY, ret] = \
                self.SapModel.Results.JointDrifts(NumberResults, Stories, Label, Names, LoadCases, StepType, StepNum,
                                                  DispX, DispY, DriftX, DriftY)

            # append all displacement results to jointdrift list
            for i in range(0, NumberResults):
                self.JointDisplacements.append((Label[i], Stories[i], LoadCases[i], DispX[i], DispY[i]))

        # set up pandas data frame and sort by drift column
        jlabels = ['label', 'Story', 'Combo', 'DispX', 'DispY']
        jdf = pd.DataFrame.from_records(self.JointDisplacements, columns=jlabels)
        story_names = jdf.Story.unique()
        # print("story names = " + str(story_names))

        # set up data frame for torsion displacement results
        tlabels = ['Story', 'Load Combo', 'Direction', 'Max Displ', 'Avg Displ', 'Ratio']
        tdf = pd.DataFrame(columns=tlabels)

        # calculate and append to df the max, avg, and ratio of story displacements in each dir.
        for dcombo in self.DriftCombos:
            for story in story_names:
                temp_df = jdf[(jdf['Story'] == story) & (jdf['Combo'] == dcombo)]
                # assume direction is X
                direction = 'X'
                averaged = abs(temp_df['DispX'].mean())
                maximumd = temp_df['DispX'].abs().max()
                averagey = abs(temp_df['DispY'].mean())

                # change direction to Y if avg y-dir displacement is higher
                if averagey > averaged:
                    averaged = averagey
                    maximumd = temp_df['DispY'].abs().max()
                    direction = 'Y'

                ratiod = maximumd / averaged

                # append info to torsion dataframe
                temp_df2 = pd.DataFrame([[story, dcombo, direction, maximumd, averaged, ratiod]], columns=tlabels)
                tdf = tdf.append(temp_df2, ignore_index=True)

        tdfSort = tdf.sort_values(by=['Ratio'], ascending=False)
        tdfSort.Ratio = tdfSort.Ratio.round(3)
        tdfSort['Max Displ'] = tdfSort['Max Displ'].round(3)
        tdfSort['Avg Displ'] = tdfSort['Avg Displ'].round(3)

        return tdfSort

    def model_close(self):
        # close the program
        ret = self.myETABSObject.ApplicationExit(False)
        self.SapModel = None
        self.myETABSObject = None
        return 'model closed'

"""
# define drift limit for DCR calc
DriftLimit = 0.01
FilePath = "C:\\Users\\Andrew-V.Young\\Desktop\\ETABS API TEST\\ETABS\\TestModel.EDB"

testModel = EtabsModel(FilePath)
drifts = testModel.story_drift_results(DriftLimit)
torsion = testModel.story_torsion_check()
print(drifts.head(10))
print("\n\n")
print(torsion)

ret = testModel.model_close()
print(ret)
"""

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# ##############################################################################################################
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
class get_model_dialog(QWidget):
    # main widget for user interface
    def __init__(self, parent=None):
        super(get_model_dialog, self).__init__(parent)
        self.initUI()
        self.chosenModel = None
        self.mess1 = ''
        self.mess_drift = ''
        self.mess_tors = ''
        self.modelPath = None
        self.driftTable = pd.DataFrame()
        self.torsTable = pd.DataFrame()

    def initUI(self):

        # add window title and prompt label text
        self.hboxtop = QHBoxLayout()
        self.setWindowTitle("Story Drift and Torsion Check Tool")
        self.layout = QVBoxLayout()  # primary window layout
        self.le = QLabel("Select ETABS model: ")
        self.hboxtop.addWidget(self.le)

        # add button to open file name, connect to open file function
        self.btn1 = QPushButton("Choose File")
        self.btn1.clicked.connect(lambda: self.getfile())
        self.hboxtop.addWidget(self.btn1)

        # add button to use open etabs model
        self.btn6 = QPushButton("Use Open Model")
        self.btn6.clicked.connect(lambda: self.activemodel())
        self.btn6.setEnabled(False) # not funtional yet
        self.hboxtop.addWidget(self.btn6)
        self.hboxtop.addStretch(1)

        self.layout.addLayout(self.hboxtop)

        # add text box to use as status notification, enter initial text
        self.statustext = QTextEdit()
        self.statustext.setText('Please use button above to choose a file')
        self.layout.addWidget(self.statustext)
        self.layout.addStretch(1)

        self.vboxdrift = QVBoxLayout()
        self.vboxtors = QVBoxLayout()
        self.hboxresults = QHBoxLayout()
        self.hboxlimit = QHBoxLayout()

        inputLimit = 0.01  # assumed limit, can be overridden by user
        # add label for drift limit
        self.lelim = QLabel()
        self.lelim.setText("Drift Limit:")
        self.hboxlimit.addWidget(self.lelim)

        # add a text input for drift limit
        self.limitText = QLineEdit()
        self.limitText.setMaxLength(6)
        self.limitText.setText(str(inputLimit))
        self.hboxlimit.addWidget(self.limitText)


        # add a button to check drift
        self.btn2 = QPushButton("Check Drift Results")
        self.btn2.clicked.connect(lambda: self.getdrift(inputLimit))
        self.vboxdrift.addWidget(self.btn2)
        self.btn2.setEnabled(False)
        self.vboxdrift.addLayout(self.hboxlimit)

        # table view to show drift results
        self.driftview = QTableView()
        self.vboxdrift.addWidget(self.driftview)

        # add a button to check torsion
        self.btn4 = QPushButton("Check Torsion Results")
        self.btn4.clicked.connect(lambda: self.gettorsion())
        self.vboxtors.addWidget(self.btn4)
        self.btn4.setEnabled(False)
        self.lespace = QLabel("")
        self.vboxtors.addWidget(self.lespace)

        # table view to show torsion check
        self.torsview = QTableView()
        self.vboxtors.addWidget(self.torsview)

        # put drift and torsion results together and into main layout
        self.hboxresults.addLayout(self.vboxdrift)
        self.hboxresults.addLayout(self.vboxtors)
        self.layout.addLayout(self.hboxresults)

        self.vboxbot = QVBoxLayout()
        self.vboxbot.addStretch(1)
        self.hboxbot = QHBoxLayout()

        # button to save results to excel file
        self.btn5 = QPushButton("Save Results")
        self.btn5.clicked.connect(lambda: self.saveResults())
        self.btn5.setEnabled(False)
        self.hboxbot.addWidget(self.btn5)
        self.hboxbot.addStretch(1)

        # button to close model
        self.btn3 = QPushButton("Close ETABS Model")
        self.btn3.clicked.connect(lambda: self.closeModel())
        self.btn3.setEnabled(False)
        self.hboxbot.addWidget(self.btn3)

        self.vboxbot.addLayout(self.hboxbot)
        self.layout.addLayout(self.vboxbot)

        self.setLayout(self.layout)

    def getfile(self):
        # function that pulls up get open file name window to select etabs model
        self.closeModel()  # close any opened model

        # open window and extract file name from outputs
        dlg = QFileDialog(caption="Choose File")
        dlg.setFileMode(QFileDialog.ExistingFile)
        dlg.setNameFilters(["ETABS Models (*.EDB)", "Backup Models (*.ebk *.$et)"])
        dlg.selectNameFilter("ETABS Models (*.EDB)")

        if dlg.exec():
            fileNames = dlg.selectedFiles()
            fileName = fileNames[0]

        # run reformatting if file chosen, otherwise no action
        if fileName:
            self.mess1 = "Selected File: \n %s \n\n" % fileName
            self.chosenModel = EtabsModel(fileName)
            self.modelPath = self.chosenModel.modelPath
            self.statustext.setText(self.mess1)
            self.btn2.setEnabled(True)
            self.btn3.setEnabled(True)
            self.btn4.setEnabled(True)
            self.btn5.setEnabled(True)
        else:
            not_opened = "No file was opened"
            self.statustext.setText(not_opened)

    def activemodel(self):
        # function to open active model - NOT WORKING YET
        self.chosenModel = EtabsModel(existinstance=True)
        self.modelPath = self.chosenModel.modelPath
        self.mess1 = "Selected open model: \n %s \n\n" % self.chosenModel.modelName
        self.statustext.setText(self.mess1)

    def getdrift(self, dlimit):
        # function that populates drift results table
        currentLimit = self.limitText.text()  # read text input
        is_lim_num = is_number(currentLimit)  # check if input is a number

        if is_lim_num:
            numberLimit = float(currentLimit)  # make input a number
            self.driftTable = self.chosenModel.story_drift_results(numberLimit)  # get drift results as dataframe
            self.mess_drift = "showing drift results for limit " + currentLimit + "\n"
            self.statustext.setText(self.mess1 + self.mess_drift + self.mess_tors)
            driftModel = pandasModel(self.driftTable)  # transform dataframe to use in QTableView
            self.driftview.setModel(driftModel)  # populate QTableView with drift results

            wmin = self.driftview.verticalHeader().width() + 24
            for i in range(driftModel.columnCount()):
                wmin += self.driftview.columnWidth(i)
            self.driftview.setMinimumWidth(wmin)
            # self.driftview.resize(800, 600)
            # https: // stackoverflow.com / questions / 41542934 / pyqt - qtablewidget - remove - scrollbar - to - show - full - table
        else:
            not_float = "Please input a number for drift limit"
            self.statustext.setText(not_float)

    def gettorsion(self):
        # function that populates torsion results table
        self.torsTable = self.chosenModel.story_torsion_check()
        self.mess_tors = "showing torsion results"
        self.statustext.setText(self.mess1 + self.mess_drift + self.mess_tors)
        torsModel = pandasModel(self.torsTable)
        self.torsview.setModel(torsModel)

        wtmin = self.torsview.verticalHeader().width() + 24
        for i in range(torsModel.columnCount()):
            wtmin += self.torsview.columnWidth(i)
        self.torsview.setMinimumWidth(wtmin)
        # https: // stackoverflow.com / questions / 41542934 / pyqt - qtablewidget - remove - scrollbar - to - show - full - table

    def saveResults(self):
        if self.modelPath:
            excel_path = self.modelPath + '/results.xlsx'
            with pd.ExcelWriter(excel_path) as writer:
                self.driftTable.to_excel(writer, sheet_name='drift_results')
                self.torsTable.to_excel(writer, sheet_name='torsion_results')
            mess_save = 'results saved to excel'
        else:
            mess_save = 'ETABS file path not found'

        self.statustext.setText(mess_save)

    # function to close current ETABS model
    def closeModel(self):
        close_mess = "nothing found to close"
        if self.chosenModel:
            close_mess = self.chosenModel.model_close()
        self.chosenModel = None
        self.statustext.setText(close_mess)

    # override close event to close etabs model first
    def closeEvent(self, event):
        # do stuff
        if self.chosenModel:
            self.closeModel()
        event.accept()  # let the window close



# set up main application with get_file_dialog widget
def main():
   app = QApplication(sys.argv)
   ex = get_model_dialog()
   ex.show()
   sys.exit(app.exec_())

# run widget
if __name__ == '__main__':
    main()