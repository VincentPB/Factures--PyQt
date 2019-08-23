
#=============================== IMPORTS =================================#

import os
import sys
import csv
import string
import pandas as pd
import datetime
if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from xlwt import Workbook
from xlrd import open_workbook
from functools import partial
from tkinter import filedialog

#============= LIST ALL SUB-DIRECTORIES ==============#

def subDir(directory):
    for elt in os.walk(directory):
        subs=elt[1]
        break
    return(subs)

#=================== LIST OF RELEVANT DATA ====================#

def listNUMAndCODE(directory):       
    DATAtemp = []
    for filename in os.listdir(directory):
        if filename.endswith('.PDF'):
            NAME = filename
            NUM = filename[22:30]
            CODE = filename[7:13]
            DATE = (filename[36:38] + '/' + filename[34:36] + '/' + filename[30:34])
            LOCATION = directory+'/'+filename
            DATAtemp.append([NAME, NUM, CODE, DATE, LOCATION])
    
    return(DATAtemp)

def listNUMAndCODEFile(filename, lastFolder):
    DATAtemp = []
    NAME = filename
    NUM = filename[22:30]
    CODE = filename[7:13]
    DATE = (filename[36:38] + '/' + filename[34:36] + '/' + filename[30:34])
    LOCATION = lastFolder[:23] + '/' + lastFolder[24:28] + '/' + lastFolder[29:31] + '/' + filename
    print(LOCATION)
    DATAtemp = [NAME, NUM, CODE, DATE, LOCATION]
    return(DATAtemp)

#==================== FILES COUNTING =====================#

def countPDF(directory):
    subs = subDir(directory)
    count = 0
    for sub in subs:
        undersubs = subDir(directory + '/' + sub)
        for undersub in undersubs:
            onlyfiles = next(os.walk(directory + '/' + sub + '/' + undersub))[2]
            count += len(onlyfiles)
    return count

#==================== DATA EXTRACTION =====================#

def extract(directory, directoryB):
    DATA = []
    subs = subDir(directory)
    subsB = subDir(directoryB)
    for sub in subs:
        undersubs = subDir(directory + '/' + sub)
        for undersub in undersubs:
            DATA += listNUMAndCODE(directory + '/' + sub + '/' + undersub)
            
    for subB in subsB:
        undersubsB = subDir(directoryB + '/' + subB)
        for undersubB in undersubsB:
            DATA += listNUMAndCODE(directoryB + '/' + subB + '/' + undersubB)    

    return(pd.DataFrame(DATA, columns=['NAME', 'NUMBILL', 'CLIENT', 'DATE', 'PATH']))


#==================== TIME CONVERTER =====================#

def dateToDateTime(date):
    dateTime = datetime.date(int(date[-4:]),int(date[3:5]),int(date[0:2]))
    return dateTime

#==================== DATA SORTING =====================#

def Sort(dateStart, dateEnd, client, output):
    listSort = []
    now = datetime.datetime.now()

    if(dateStart=='' or dateStart=="JJ/MM/AAAA"):
        dateStartTime = datetime.date(1,1,1)
    else:
        dateStartTime = dateToDateTime(dateStart)
        
    if(dateEnd=='' or dateEnd=="JJ/MM/AAAA"):
        dateEndTime = datetime.date(now.year,now.month,now.day)
    else:
        dateEndTime = dateToDateTime(dateEnd)

    L = len(output)
    for i in range(L):
        dateF = datetime.date(int(output.DATE[i][6:10]), int(output.DATE[i][3:5]), int(output.DATE[i][0:2]))
        clientF = output.CLIENT[i]
                
        if(client==''):
            if((dateF >= dateStartTime) and (dateF <= dateEndTime)):
                listSort.append(output.NAME[i])
                
        else:
            if((dateF >= dateStartTime) and (dateF <= dateEndTime) and (client==clientF)):
                listSort.append(output.NAME[i])
                    
    return listSort


def factSort(fact, listSort): #Sort by invoice number

    if (fact==''):
        return listSort

    listSortFinal=[]
    
    for i in listSort:
        factF = i[22:29]
        if(fact==factF):
            listSortFinal.append(i)
             
    return listSortFinal

#==================== LIST FILES =====================#

def listPDF(directory):       
    pdfFiles = []
    for filename in os.listdir(directory):
        if filename.endswith('.PDF'):
            pdfFiles.append(filename)
    return(pdfFiles)

def getfiles(dirpath): #Get recent files
    a=os.listdir(dirpath)
    return a


directory = r'\\srv-FIc\Archivage_factures_srv_map'
directoryB = r'\\10.202.72.22\Factures'
outputFile = 'DATA.csv'
print("LENGTH CALCULATION PROCESSING, PLEASE WAIT...")
nbPDF = countPDF(directory) + countPDF(directoryB)
print("LENGTH CALCULATION DONE")
output = pd.read_csv(outputFile)
LO = len(output)

if (LO == 0):
    print('EMPTY FILE')
    DATAF = extract(directory, directoryB)  
    DATAF.to_csv(outputFile, sep = ',', index=False)
    output = pd.read_csv(outputFile)
    
if (LO != nbPDF):
    print('FILES TO ADD : ', abs(LO-nbPDF))
    Year = [f.name for f in os.scandir(directoryB) if f.is_dir() ][-1]
    NumDos = [f.name for f in os.scandir(directoryB+'/'+Year) if f.is_dir() ][-1]
    lastFolder = directoryB+'/'+Year+'/'+NumDos
    lastFilerPDF = getfiles(lastFolder)[-(nbPDF - LO):]
    with open(outputFile, 'a') as f:
        writer = csv.writer(f, lineterminator='\n')
        for i in lastFilerPDF:
            newLine = listNUMAndCODEFile(i, lastFolder)  
            writer.writerow(newLine)

else:
    print('NO FILES NO ADD')

print('DATABASE UPDATED')  
output = pd.read_csv(outputFile)

#=========================================#

def filterD(line1, line2, line3, line4, listeSel):

    listeSel.clear()
    
    dateStart = line1.text()
    dateEnd = line2.text()
    client = line3.text()
    fact = line4.text()

    listeFiltre0 = Sort(dateStart, dateEnd, client, output)
    listeFiltre = factSort(fact, listeFiltre0)

    for i in range(len(listeFiltre)):
        #listSel.insert(i, listeFiltre[i][:-4])
        item = QListWidgetItem(listeFiltre[i][:-4])
        listeSel.addItem(item)

#=========================== DISPLAY FUNCTION ============================#

def showDialog(): #PopUp de fin de traitement
    msgBox = QMessageBox()
    msgBox.setGeometry(500,350, 200, 200)
    msgBox.setText("<p align='center'>Le dédoublonnage a été effectué avec succès </p>")
    msgBox.setWindowTitle("Traitement terminé")
    msgBox.setFont(QFont("Calibri", 11, QFont.Bold))
    msgBox.setStyleSheet(
    "QPushButton {"
    " font: bold 14px;"
    " min-width: 10em;"
    " padding: 3px;"
    " margin-right:4.5em;"
    "}"
    "* {"
    " margin-right:1.8em;"
    "min-width: 22em;"
    "}"
    );
    msgBox.exec()

def aProposDe(): #PopUp 'A propos'
    msgBox = QMessageBox()
    msgBox.setGeometry(510,350, 200, 200)
    msgBox.setText("<p align='center'>Cette application est une propriété</p> \n <p align='center'>Stela Produits Pétroliers</p>")
    msgBox.setWindowTitle("A propos")
    msgBox.setFont(QFont("Calibri", 11, QFont.Bold))
    msgBox.setStyleSheet(
    "QPushButton {"
    " font: bold 14px;"
    " min-width: 10em;"
    " padding: 3px;"
    " margin-right:3.5em;"
    "}"
    "* {"
    " margin-right:1.8em;"
    "min-width: 20em;"
    "}"
    );
    msgBox.exec()

def openFileNameDialog(): #Retourne le nom du fichier sélectionné
        fileName = QFileDialog.getOpenFileName()
        return fileName[0]

class MyMainWindow(QMainWindow): #Fenêtre

    def __init__(self, parent=None):

        super(MyMainWindow, self).__init__(parent)
        self.form_widget = Example(self) 
        self.setCentralWidget(self.form_widget)
        self.setGeometry(450, 250, 850, 650)
        self.setWindowTitle('Dédoublonnage')
        self.setWindowIcon(QIcon('stela.ico'))

        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu('Options')
        
        exitButton = QAction(QIcon('exit24.png'), 'Quitter', self)
        exitButton.setShortcut('Ctrl+Q')
        exitButton.setStatusTip("Quitter l'application")
        exitButton.triggered.connect(self.close)
        aPropos = QAction(QIcon('exit24.png'), 'A propos', self)
        aPropos.triggered.connect(aProposDe)

        fileMenu.addAction(aPropos)
        fileMenu.addAction(exitButton)

        self.showMaximized()

class Example(QWidget): #Widget
    
    def __init__(self, parent):
        super(Example, self).__init__(parent)
        self.initUI()
        
    def initUI(self):

        buttonT = QPushButton('RECHERCHER', self)
        buttonT.setToolTip('Chercher les factures correspondantes')
        buttonT.clicked.connect(lambda : filterD(self.line1, self.line2, self.line3, self.line4, self.listWidget))
        buttonT.move(1150, 15)
        buttonT.setFont(QFont("Calibri", 11, QFont.Bold))
        buttonT.resize(100, 40)
        
        self.nameLabel1 = QLabel(self)
        self.nameLabel1.setText('<p align=center>Date Début<br>(JJ/MM/AAAA)</p>')
        self.nameLabel1.setFont(QFont("Calibri", 11, QFont.Bold))
        self.line1 = QLineEdit(self)
        self.line1.setFont(QFont("Calibri", 11, QFont.Bold))
        self.nameLabel2 = QLabel(self)
        self.nameLabel2.setText('<p align=center>Date Fin<br>(JJ/MM/AAAA)</p>')
        self.nameLabel2.setFont(QFont("Calibri", 11, QFont.Bold))
        self.line2 = QLineEdit(self)
        self.line2.setFont(QFont("Calibri", 11, QFont.Bold))
        self.nameLabel3 = QLabel(self)
        self.nameLabel3.setText('<p align=center>N° Client<br>(6 chars)</p>')
        self.nameLabel3.setFont(QFont("Calibri", 11, QFont.Bold))
        self.line3 = QLineEdit(self)
        self.line3.setFont(QFont("Calibri", 11, QFont.Bold))
        self.nameLabel4 = QLabel(self)
        self.nameLabel4.setText('<p align=center>N° Facture<br>(7 chars)</p>')
        self.nameLabel4.setFont(QFont("Calibri", 11, QFont.Bold))
        self.line4 = QLineEdit(self)
        self.line4.setFont(QFont("Calibri", 11, QFont.Bold))

        self.line1.move(180, 20)
        self.line1.resize(130, 28)
        self.line2.move(450, 20)
        self.line2.resize(130, 28)
        self.line3.move(720, 20)
        self.line3.resize(130, 28)
        self.line4.move(990, 20)
        self.line4.resize(130, 28)
    
        self.nameLabel1.move(80, 18)
        self.nameLabel2.move(350, 18)
        self.nameLabel3.move(620, 18)
        self.nameLabel4.move(890, 18)

        self.listWidget = QListWidget()
        self.listWidget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.listWidget.setGeometry(QRect(10, 10, 211, 291))
        self.listWidget.setFont(QFont("Calibri", 12))

        self.BLANC = QLabel(self)
        self.BLANC.setText('                        ')
        self.BLANC2 = QLabel(self)
        self.BLANC2.setText('                        ')
        
        #○self.listWidget.itemClicked.connect(lambda: print('BONSOIR'))


        buttonO = QPushButton('OUVRIR', self)
        buttonO.setToolTip('Ouvrir les factures sélectionnées')
        buttonO.clicked.connect(lambda : ouverture(self.listWidget))
        buttonO.setFont(QFont("Calibri", 11, QFont.Bold))
        buttonO.resize(100, 40)

        self.layout = QGridLayout(self)
        self.layout.addWidget(buttonT, 0, 9)
        self.layout.addWidget(self.nameLabel1, 0, 1)
        self.layout.addWidget(self.nameLabel2, 0, 3)
        self.layout.addWidget(self.nameLabel3, 0, 5)
        self.layout.addWidget(self.nameLabel4, 0, 7)
        self.layout.addWidget(self.line1, 0, 2)
        self.layout.addWidget(self.line2, 0, 4)
        self.layout.addWidget(self.line3, 0, 6)
        self.layout.addWidget(self.line4, 0, 8)
        self.layout.addWidget(self.BLANC, 2,5)
        self.layout.addWidget(self.listWidget, 3, 4, 1, 3)
        self.layout.addWidget(buttonO, 3, 8)
        
        self.layout.addWidget(self.BLANC2, 4,5)
        

        self.setLayout(self.layout)
 
        self.show()

#================================ DISPLAY =================================#

def ouverture(listSel):
    	 
    items = listSel.selectedItems()
    x = []
    for i in range(len(items)):
        x.append(str(listSel.selectedItems()[i].text()))
    for j in range (len(x)):
        numF = str(x[j])[22:29]
        i=0
        while(str(numF))[1:]!=str(output.NUMBILL[i]):
            i+=1
        if(output.PATH[i][4]=='v'):
            realPath = output.PATH[i][:37]+'\\'+output.PATH[i][37:42]+'\\'+output.PATH[i][42:45]+'\\'+output.PATH[i][45:]
        else:
            realPath = output.PATH[i][:24]+'\\'+output.PATH[i][24:29]+'\\'+output.PATH[i][29:32]+'\\'+output.PATH[i][32:]
        os.startfile(realPath)

        
app = QApplication([])

        #----------------- STYLE DARK ----------------------#

app.setStyle('Fusion')  
dark_palette = QPalette()
dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
dark_palette.setColor(QPalette.WindowText, Qt.white)
dark_palette.setColor(QPalette.Base, QColor(25, 25, 25))
dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
dark_palette.setColor(QPalette.ToolTipBase, Qt.white)
dark_palette.setColor(QPalette.ToolTipText, Qt.white)
dark_palette.setColor(QPalette.Text, Qt.white)
dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
dark_palette.setColor(QPalette.ButtonText, Qt.white)
dark_palette.setColor(QPalette.BrightText, Qt.red)
dark_palette.setColor(QPalette.Link, QColor(42, 130, 218))
dark_palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
dark_palette.setColor(QPalette.HighlightedText, Qt.black)
app.setPalette(dark_palette)
app.setStyleSheet("QToolTip { color: #ffffff; background-color: #2a82da; border: 1px solid white; }")

        #----------------- AFFICHAGE ----------------------#

foo = MyMainWindow()
foo.show()
sys.exit(app.exec_())
