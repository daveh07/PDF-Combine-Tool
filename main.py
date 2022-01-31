#######################################################################################################################
## PDF COMBINE - DESKTOP UTILITY
## BY: DAVID HILL
## PROJECT MADE WITH: Qt Designer and PySide6
## V: 1.0.0
## 31/01/2022
#######################################################################################################################

# IMPORT PACKAGES
import sys, os, io, platform, glob
from os import path

from PyPDF2 import PdfFileMerger

if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ';' + os.environ['PATH']

# IMPORT PYSIDE6 MODULES & LIBRARIES
from PySide6 import QtCore
from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QMainWindow, QListWidget, QListWidgetItem, QAbstractItemView, QLineEdit, \
    QFileDialog, QMessageBox, QSizeGrip
import PyPDF2

# IMPORT QT DESIGNER GUI
from ui_main import Ui_MainWindow


# Function to define teh resource path
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)


# class ListWidget(QListWidget):
#     def __init__(self, parent=None):
#         super().__init__(parent=None)
#         self.setAcceptDrops(True)
#         self.setDragDropMode(QAbstractItemView.InternalMove)
#         self.setSelection(QAbstractItemView.ExtendedSelection)
#
#     def dragEnterEvent(self, event):
#         if event.mimeData().hasUrls():
#             event.accept()
#         else:
#             return super().dragEnterEvent(event)
#
#     def dragMoveEvent(self, event):
#         if event.mimeData().hasUrls():
#             event.setDropAction(Qt.CopyAction)
#             event.accept()
#         else:
#             super().dragMoveEvent(event)
#
#     def dropEvent(self, event):
#         if event.mimeData().hasUrls():
#             event.setDropAction(Qt.CopyAction)
#             event.accept()
#
#             file_list = []
#
#             for url in event.mimeData().urls():
#                 if url.isLocalFile():
#                     if url.tostring().endswith(".pdf"):
#                         file_list.append(str(url.toLocalFile()))
#                 self.addItems(file_list)
#             else:
#                 return super().dropEvent(event)

# --------------------------------------------- QLine Edit Output Field ----------------------------------------------- #
# class output_field(QLineEdit):
#     def __init__(self):
#         super().__init__()
#
#     def dragEnterEvent(self, event):
#         if event.mimeData().hasUrls():
#             event.accept()
#         else:
#             self.ignore()
#
#     def dragMoveEvent(self, event):
#         if event.mimeData().hasUrls():
#             event.setDropAction(Qt.CopyAction)
#             event.accept()
#         else:
#             event.ignore()
#
#     def dropEvent(self, event):
#         if event.mimeData().hasUrls():
#             event.setDropAction(Qt.CopyAction)
#             event.accept()
#
#             if event.mimeData().urls():
#                 event.setText(event.mimeData().urls()[0].toLocalFile())
#             else:
#                 event.ignore()


# MAIN WINDOW SET-UP
class EziPDF_App(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.outputText = self.ui.lineEdit

        self.setWindowTitle("EziPDF - Combine Utility")

        all_files = []
        for dirpath, dirnames, filenames in os.walk("."):
            for filename in [f for f in filenames if f.endswith(".pdf")]:
                all_files.append(os.path.join(dirpath, filename))



        def moveWindow(event):
            # IF MAXIMIZED CHANGE TO NORMAL
            if self.returnStatus() == 1:
                self.maximize_restore()

        # --------------------------------------------- REMOVE DEFAULT TITLE BAR ---------------------------------------------- #
        self.setWindowFlag(QtCore.Qt.WindowType.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground)

        # MAX/MIN & CLOSE APPLICATION BUTTONS
        self.ui.closeBtn.clicked.connect(self.close)
        self.ui.minimizeBtn.clicked.connect(self.showMinimized)
        self.ui.restoreBtn.clicked.connect(self.maximize_restore)

        # WINDOW SIZE GRIP

        QSizeGrip(self.ui.size_grip)

        self.upload_path = "PDF_TEST"

# --------------------------------------- QLISTWIDGET COMBINE PDF FUNCTIONS ------------------------------------------ #
        #        ListWidget()
        self.list = self.ui.listWidget
        self.list.setSpacing(2)
        self.list.setDragDropMode(self.list.InternalMove)
        #        self.list.setSelectionMode(self.list.MultiSelection)

        self.combine_pdf = self.ui.pushButton_6.clicked.connect(self.mergePDF)
        self.saveFile_Name = self.ui.pushButton_2.clicked.connect(self.setFileName)
        self.upload_files = self.ui.pushButton_3.clicked.connect(self.uploadBtn)
        self.clear_list = self.ui.pushButton_7.clicked.connect(self.clearList)
        self.remove_item = self.ui.pushButton_4.clicked.connect(self.removeListItem)
        self.move_item_up = self.ui.pushButton_5.clicked.connect(self.moveItemUp)
        self.move_item_down = self.ui.pushButton.clicked.connect(self.moveItemDown)

        self.filesList = self.listFiles()

# --------------------------------------- RESTORE/MAXIMISE WINDOW FUNCTION #------------------------------------------ #
    def maximize_restore(self):
        global GLOBAL_STATE
        GLOBAL_STATE = int(self.windowState())
        status = GLOBAL_STATE

        if status == 0:
            self.showMaximized()
            GLOBAL_STATE = 1

        else:
            GLOBAL_STATE = 0
            self.showNormal()

    ## ==> RETURN STATUS
    def returnStatus(self):
        return GLOBAL_STATE

    ## ==> SET STATUS
    def setStatus(status):
        global GLOBAL_STATE
        GLOBAL_STATE = status

# -------------------------------------------- Window Header Functions ----------------------------------------------- #
    def mousePressEvent(self, event):
        self.oldPosition = event.globalPos()

    def mouseMoveEvent(self, event):
        delta = QPoint(event.globalPos() - self.oldPosition)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.oldPosition = event.globalPos()

    def setFileName(self):
        path = self._getSaveFilePath()
        if path:
            self.outputText.setText('')

    def dialogMessage(self, message):
        dlg = QMessageBox(self)
        dlg.setWindowTitle('EziPDF Merger')
        dlg.setIcon(QMessageBox.Information)
        dlg.setText(str(message))
        dlg.show()

    def _getSaveFilePath(self):
        file_save_path, _ = QFileDialog.getSaveFileName(self, 'Save PDF File', os.getcwd(), 'PDF file (*.pdf)')
        return file_save_path

    def mergePDF(self):
        if not self.ui.lineEdit.text():
            self.dialogMessage('Save File to:')
            self.setFileName()
            return

        if self.List.count() > 0:
            pdfMerger = PdfFileMerger()
            try:
                for i in range(self.List.count()):
                    pdfMerger.append(self.List.item(i).text())

                pdfMerger.write(self.outputFile.text())
                self.dialogMessage('PDF Merge Complete')

            except Exception as e:
                self.dialogMessage(e)

        else:
            self.dialogMessage('Queue is empty')

# -------------------------------------------- BUTTON FUNCTIONS ------------------------------------------------------ #
    # -- Upload Button -- #
    def uploadBtn(self):
        dialog = QFileDialog()
        folder_path = dialog.getExistingDirectory(None, "Select Folder")
        return folder_path

    # -- Clear Button -- #
    def clearList(self):
        self.list.clear()

    # -- Remove Button -- #
    def removeListItem(self):
        for item in self.list.selectedItems():
            self.list.takeItem(self.list.row(item))

    # -- Move Item Buttons -- #
    def moveItemUp(self):
        currentRow = self.list.currentRow()
        currentItem = self.list.takeItem(currentRow)
        self.list.insertItem(currentRow - 1, currentItem)

    def moveItemDown(self):
        currentRow = self.list.currentRow()
        currentItem = self.list.takeItem(currentRow)
        self.list.insertItem(currentRow + 1, currentItem)

# -------------------------------------------- List files in QListWidget --------------------------------------------- #
    def listFiles(self):

        fileList = []

        fileNames = os.listdir(self.upload_path)
        for file in fileNames:
            fileList.append(file)
            self.list.addItem(file)

        return fileList


# ------------------------------------------------ RUN APPLICATION --------------------------------------------------- #
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("fusion")
    window = EziPDF_App()
    window.show()
    sys.exit(app.exec())
