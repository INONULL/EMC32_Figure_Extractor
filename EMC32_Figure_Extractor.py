#encoding: euc-kr

#EMC32 Figure Exporter GUI
import os, sys, locale, webbrowser
from PyQt6.QtWidgets import QDialog, QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox, QProgressBar, QCheckBox, QMenuBar, QLineEdit, QHBoxLayout
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QCursor, QFont, QIcon, QAction
from wand.image import Image as WandImage

magick_home = os.getcwd() + os.sep
os.environ["PATH"] += os.pathsep + magick_home + os.sep
os.environ["MAGICK_HOME"] = magick_home
os.environ["MAGICK_CODER_MODULE_PATH"] = magick_home + "modules" + os.sep + "coders"

class Chkbox(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.tickedFile = []
        self.tickedFolder = []
        self.checkboxLineEditFolderName = []
        self.checkboxLineEditFileName = []
        self.insert_option = False
        self.setWindowTitle("Choose file name(s), which you want to extract!")
        self.layout = QVBoxLayout()
        self.filename_label = QLabel("If you want to rename file(s), type in!\n")
        self.layout.addWidget(self.filename_label)


        self.folder_starts_here = QLabel("=================  Rename [folder]  ===============\n")
        self.layout.addWidget(self.folder_starts_here)
        self.insert_option = QCheckBox("Insert", self)
        self.insert_option.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.insert_option.stateChanged.connect(self.on_insert_option)
        self.layout.addWidget(self.insert_option)
        for folder in EMC32Extract.folder_list_mod:
            self.hlayout_Chkbox = QHBoxLayout()  # Create a new QHBoxLayout for each file
            checkboxfolder = QCheckBox(folder, self)
            self.tickedFolder.append(checkboxfolder)
            checkboxLineEditFolderName = QLineEdit(self)
            self.checkboxLineEditFolderName.append(checkboxLineEditFolderName)
            self.hlayout_Chkbox.addWidget(checkboxfolder)
            self.hlayout_Chkbox.addWidget(checkboxLineEditFolderName)
            self.layout.addLayout(self.hlayout_Chkbox)  # Add the QHBoxLayout to the main QVBoxLayout


        self.file_starts_here = QLabel("\n================  Rename [file]  ===============\n")
        self.layout.addWidget(self.file_starts_here)
        for file in EMC32Extract.file_list:
            self.hlayout_Chkbox = QHBoxLayout()  # Create a new QHBoxLayout for each file
            checkbox = QCheckBox(file, self)
            self.tickedFile.append(checkbox)
            checkboxLineEditName = QLineEdit(self)
            self.checkboxLineEditFileName.append(checkboxLineEditName)
            self.hlayout_Chkbox.addWidget(checkbox)
            self.hlayout_Chkbox.addWidget(checkboxLineEditName)
            self.layout.addLayout(self.hlayout_Chkbox)  # Add the QHBoxLayout to the main QVBoxLayout

        self.dpi_size_label = QLabel("\nDefault DPI='300'\nDPI:")
        self.dpi_size = QLineEdit(self)
        self.layout.addWidget(self.dpi_size_label)
        self.layout.addWidget(self.dpi_size)
        self.button_ok = QPushButton("OK")
        self.button_cancel = QPushButton("Cancel")
        self.button_ok.clicked.connect(self.accept)
        self.button_cancel.clicked.connect(self.reject)

        self.layout.addWidget(self.button_ok)
        self.layout.addWidget(self.button_cancel)
        self.adjustSize()
        self.setLayout(self.layout)
    def on_insert_option(self, state):
        self.insert_option = state == Qt.CheckState.Checked.value
        return self.insert_option
 
class EMC32Extractor(QWidget):
    def __init__(self):
        super().__init__()
        self.input_folder = ""
        self.output_folder = ""
        self.file_list = []
        self.folder_list = []
        self.folder_list_mod = []
        self.button_press_count_folder = 0
        self.setWindowIcon(QIcon(os.getcwd() + os.sep + 'icon.ico'))   
        self.init_ui()

    def init_ui(self): 
        self.setWindowFlags(Qt.WindowType.WindowCloseButtonHint | Qt.WindowType.WindowMinimizeButtonHint | Qt.WindowType.Window)
        self.setWindowTitle("EMC32_Graph_Extractor")
        # check_locale_button = QPushButton("Check_Locale", self)
        # check_locale_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        # check_locale_button.clicked.connect(self.on_check_locale)
        input_folder_button = QPushButton("Select Input Folder", self)
        input_folder_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        input_folder_button.clicked.connect(self.on_select_input_folder)
        self.input_folder_label = QLabel("Input Folder:", self)
        self.checked_file_label = QLabel("Target(s)(FileName):\n[]", self)
        self.checked_changed_folder_label = QLabel("Rename(FolderName):\n[]", self)
        self.checked_changed_file_label = QLabel("Rename(FileName):\n[]", self)

        output_folder_button = QPushButton("Select Output Folder", self)
        output_folder_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        output_folder_button.clicked.connect(self.on_select_output_folder)
        self.output_folder_label = QLabel("Output Folder: ", self)

        extract_button = QPushButton("Extract", self)
        extract_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        extract_button.clicked.connect(self.extract_image)

        self.status_label = QLabel("", self)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)

        menubar = QMenuBar(self)
        help_menu = menubar.addMenu("Help")

        howtouse_action = QAction("How To Use", self)
        howtouse_action.triggered.connect(self.show_howtouse_dialog)
        help_menu.addAction(howtouse_action)

        reference_action = QAction("Library Used", self)
        reference_action.triggered.connect(self.show_reference_dialog)
        help_menu.addAction(reference_action)

        algorithm_action = QAction("Algorithm Used", self)
        algorithm_action.triggered.connect(self.show_algorithm_dialog)
        help_menu.addAction(algorithm_action)

        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

        vbox = QVBoxLayout()
        #vbox.addWidget(check_locale_button)
        vbox.addWidget(input_folder_button)
        vbox.addWidget(self.input_folder_label)
        vbox.addWidget(self.checked_file_label)
        vbox.addWidget(self.checked_changed_folder_label)
        vbox.addWidget(self.checked_changed_file_label)
        vbox.addWidget(output_folder_button)
        vbox.addWidget(self.output_folder_label)
        vbox.addWidget(extract_button)
        vbox.addWidget(self.status_label)
        vbox.addWidget(self.progress_bar)
        vbox.setMenuBar(menubar)
        self.setLayout(vbox)
        self.setMinimumSize(500, 300)
        self.adjustSize()
        self.setMaximumSize(1000, 450)
        vbox = QVBoxLayout()

        self.show()

    def show_howtouse_dialog(self):
        pdfpath = os.path.realpath(os.getcwd() + os.sep + 'HowToUse.pdf')
        if os.path.isfile(pdfpath):
            webbrowser.open(pdfpath)
        
    def show_about_dialog(self):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("About EMC32_Figure_Exporter")
        msg_box.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        msg_box.setText("This is 'EMC32_Figure_Extraction' application.\nCoded by JunSungLEE \nContact(Bug_Report): ljs_fr@nate.com")
        msg_box.exec()

    def show_reference_dialog(self):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Library Used")
        msg_box.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        msg_box.setText("Python3.11.5(Wand, PyQt6, Pyinstaller); ImageMagick(Wand)")
        msg_box.exec()

    def show_algorithm_dialog(self):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Algorithm Used")
        msg_box.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        msg_box.setText("[Python Libray(Wand)], ImageMagick\nWand(Image(WMF -> PNG))")
        msg_box.exec()


    # def on_check_locale(self): #Check Current Locale.
    #     try:                
    #         msg_box = QMessageBox(self)
    #         msg_box.setWindowTitle("Current System Locale")
    #         msg_box.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
    #         current_locale = locale.getlocale()[0]
    #         if not current_locale == 'en_US':
    #             msg_box.setText('Current_Locale: ' + ''.join(current_locale)+'\n Locale will be changed to (en-US)')
    #             locale.setlocale(locale.LC_ALL, 'en_US')
    #             current_locale = locale.getlocale()
    #             msg_box.exec()
    #         else:
    #             msg_box.setText('Current_Locale: ' + ''.join(current_locale)+'\n You are Good to GO!')
    #             msg_box.exec()
    #     except FileNotFoundError:
    #         msg_box.setText("Error: Locale registry key not found.")
    #         msg_box.exec()
    #     except Exception as e:
    #         msg_box.setText("Error:", e)
    #         msg_box.exec()

    def on_select_input_folder(self):
        self.num_images = 0
        self.file_name = []
        self.folder_name = []
        self.tickedFile = []
        self.tickedFolder = []
        self.file_list = []
        self.folder_list_mod = []
        self.checkboxLineEditFileName = []
        self.checkboxLineEditFolderName = []
        self.checkbox_insert_option_list = []
        default_folder = os.path.expanduser("~")
#여기 밑 부터 Checkboxfolder_mod 추가 하면 됨
        if self.button_press_count_folder == 0:
            folder = QFileDialog.getExistingDirectory(self, "Select Input Folder", default_folder)
        else:
            folder = QFileDialog.getExistingDirectory(self, "Select Input Folder")
        if folder:
            self.input_folder = folder
            self.input_folder_label.setText(f"Input Folder: {self.input_folder}")
            self.button_press_count_folder = 1

            for folder_name in os.listdir(self.input_folder):
                folder_path = os.path.join(self.input_folder, folder_name)
                folder_graph_path = os.path.join(folder_path, 'Graphics')  # Corrected path creation
                folder_name_mod = folder_name[:5].replace(".","_") + folder_name[5:]
                folder_name_mod = folder_name_mod.split("_")
                if len(folder_name_mod) <= 1:
                    msg_box = QMessageBox(self)
                    msg_box.setText('Wrong folder selection')
                    msg_box.exec()
                    return
                folder_name_mod = folder_name_mod[1]
                self.folder_list.append(folder_name)
                self.folder_list_mod.append(folder_name_mod)

                if os.path.isdir(folder_graph_path):
                    for filename in os.listdir(folder_graph_path):
                        if filename.lower().endswith('.wmf'):
                            self.file_list.append(filename.replace('.wmf',""))

            self.folder_list = list(set(self.folder_list)) 
            self.folder_list_mod = list(set(self.folder_list_mod))
            self.file_list = list(set(self.file_list))

            if not self.file_list:
                QMessageBox.warning(self, "No WMF file detected", "Please select valid directory")
                return
            else:
                self.chkbox = Chkbox(self)
                result = self.chkbox.exec()

                if self.chkbox.dpi_size.text().isdigit():
                    self.dpi_size = self.chkbox.dpi_size.text()
                elif self.chkbox.dpi_size.text()=="":
                     self.dpi_size = "300"
                else:
                    QMessageBox.warning(self,"Enter Dpi_size again", "Please enter the value correctly")
                    result = QDialog.DialogCode.Rejected 
                    return
                if result == QDialog.DialogCode.Accepted:
                    for index, checkboxfolder in enumerate(self.chkbox.tickedFolder):
                        if checkboxfolder.isChecked():
                            for folder_name in os.listdir(self.input_folder):
                                if not self.chkbox.checkboxLineEditFolderName[index].text()=="" and checkboxfolder.text() in folder_name:
                                    if self.chkbox.insert_option == True:
                                        foldername = str(folder_name[:4].replace('.','_',1) + folder_name[4:])
                                        cnt = 0
                                        while cnt < len(folder_name) and folder_name[cnt] == '0':
                                            cnt += 1
                                        self.checkbox_insert_option_list = foldername.split('_')
                                        self.checkboxLineEditFolderName.append('_'.join([self.checkbox_insert_option_list[0], self.chkbox.checkboxLineEditFolderName[index].text()]) + '_'.join(self.checkbox_insert_option_list[1:]))
                                        self.tickedFolder.append(folder_name)
                                    else:
                                        self.checkboxLineEditFolderName.append(folder_name.replace(checkboxfolder.text(), self.chkbox.checkboxLineEditFolderName[index].text()))
                                        self.tickedFolder.append(folder_name)
                                elif self.chkbox.checkboxLineEditFolderName[index].text()=="":
                                    self.checkboxLineEditFolderName.append(folder_name)
                                    self.tickedFolder.append(folder_name)
                        else:
                            for folder_name in os.listdir(self.input_folder):
                                self.checkboxLineEditFolderName.append(folder_name)
                                self.tickedFolder.append(folder_name)
                    self.checkboxLineEditFolderName = list(set(self.checkboxLineEditFolderName))
                    self.tickedFolder = list(set(self.tickedFolder))
                    for index, checkbox in enumerate(self.chkbox.tickedFile):
                        if checkbox.isChecked():
                            self.tickedFile.append(checkbox.text())
                            if not self.chkbox.checkboxLineEditFileName[index].text()=="":
                                self.checkboxLineEditFileName.append(self.chkbox.checkboxLineEditFileName[index].text())
                            else:
                                self.checkboxLineEditFileName.append(checkbox.text())

                    if self.tickedFile == []:
                        QMessageBox.warning(self, "Nothing Selected", "Please tick the file(s) you want to convert")
                        self.input_folder = ""
                        self.tickedFolder = []
                        self.tickedFile = []
                        self.checkboxLineEditFileName = []
                        self.input_folder_label.setText(f"Input Folder: {self.input_folder}")
                        self.checked_file_label.setText(f"Target(s)FileName: \n{self.tickedFile}")
                        self.checked_changed_folder_label.setText(f"Reanme(FolderName): \n{self.checkboxLineEditFolderName}")
                        self.checked_changed_file_label.setText(f"Rename(FileName): \n{self.checkboxLineEditFileName}")
                        return
                    else:
                        self.checked_file_label.setText(f"Target(s)FileName: \n{self.tickedFile}")
                        self.checked_changed_folder_label.setText(f"Rename(FolderName): \n{self.checkboxLineEditFolderName}")
                        self.checked_changed_file_label.setText(f"Rename(FileName): \n{self.checkboxLineEditFileName}")



                for folder_name in os.listdir(self.input_folder):
                    folder_path = os.path.join(self.input_folder, folder_name)
                    folder_graph_path = os.path.join(folder_path, 'Graphics')  # Corrected path creation
                    self.folder_name.append(folder_name)

                    if os.path.isdir(folder_graph_path):
                        for filename in os.listdir(folder_graph_path):
                            for name in self.tickedFile:
                                if filename.lower() == name.lower() + '.wmf':
                                    self.file_name.append(os.path.join(folder_graph_path, filename))
                                    self.num_images += 1

        else:

            Chkbox.tickedFile = []
            Chkbox.tickedFolder = []
            Chkbox.checkboxLineEditFolderName = []
            Chkbox.checkboxLineEditFileName = []
            self.file_list = []
            self.file_name = []
            self.input_folder = ""
            self.folder_list = []
            self.folder_list_mod = []
            self.tickedFile = []
            self.checkboxLineEditFileName = []
            self.checkboxLineEditFolderName = []
            self.checkbox_insert_option_list = []
            self.checked_changed_file_label = ""
            self.checked_changed_folder_label = ""
            self.tickedFolder = []
            self.input_folder_label.setText(f"Input Folder: {self.input_folder}")
            self.checked_file_label.setText(f"Target(s): \n{self.tickedFile}")


    def on_select_output_folder(self):
        default_folder = os.path.expanduser("~") # 향 후 수정
        if self.button_press_count_folder == 0:
            folder = QFileDialog.getExistingDirectory(self, "Select Output Folder", default_folder)
        else:
            folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_folder = folder
            self.output_folder_label.setText(f"Output Folder: {self.output_folder}")
        self.button_press_count_folder = 1
           
# '.'을 '_'로 치환 4번째 <- 완료
# filename 변경 기능 <- filename 기능에 대한 피드백 필요 <- 완료
    def extract_image(self):
            self.resize_thread = EMC32FigureExportThread(self.input_folder, self.output_folder, self.dpi_size, self.num_images, self.tickedFolder, self.tickedFile, self.checkboxLineEditFolderName, self.checkboxLineEditFileName)
            self.resize_thread.nofile.connect(self.show_error)
            self.resize_thread.progressChanged.connect(self.update_progress)
            self.resize_thread.finished.connect(self.extraction_finished)
            self.resize_thread.start()

    def update_progress(self, progress, processed_images, num_images):
        self.status_label.setText(f"{processed_images} / {num_images}")
        self.progress_bar.setValue(progress)

    def extraction_finished(self):
        self.status_label.setText("Done")
        def showDialog():
            msgBox = QMessageBox()
            msgBox.setWindowIcon(QIcon(os.getcwd() + os.sep + 'icon.ico'))
            msgBox.setText("Image extraction is completed\n Click 'OK' to Open Output Folder")
            msgBox.setWindowTitle("Job Finished!")
            msgBox.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
            returnValue = msgBox.exec()
            if returnValue == QMessageBox.StandardButton.Ok:
                path = os.path.realpath(self.output_folder)
                os.startfile(path)
        showDialog()
        #Reset All Variable
        Chkbox.tickedFile = []
        Chkbox.tickedFolder = []
        Chkbox.checkboxLineEditFolderName = []
        Chkbox.checkboxLineEditFileName = []
        self.file_list = []
        self.file_name = []
        self.folder_list = []
        self.folder_list_mod = []
        self.tickedFolder = []
        self.progress_bar.setValue(0)
        self.input_folder = ""
        self.output_folder = ""
        self.tickedFile = []
        self.checkboxLineEditFolderName = []
        self.checkboxLineEditFileName = []
        self.checkbox_insert_option_list = []
        self.num_images = 0
        self.progress = 0
        self.input_folder_label.setText(f"Input Folder: {self.input_folder}")
        self.checked_file_label.setText(f"Target(s): \n{self.tickedFile}")
        self.checked_changed_folder_label.setText(f"Rename(FolderName): \n{self.checkboxLineEditFolderName}")
        self.checked_changed_file_label.setText(f"Rename(FileName): \n{self.checkboxLineEditFileName}")
        self.output_folder_label.setText(f"Output Folder: {self.output_folder}")

    def show_error(self):
        QMessageBox.warning(self, "No Images Found in Subfolders", f"No images with the selected extensions found in Subfolders.")

class EMC32FigureExportThread(QThread):
    progressChanged = pyqtSignal(int, int, int)
    finished = pyqtSignal()
    nofile = pyqtSignal()

    def __init__(self, input_folder, output_folder, dpi_size, num_images, tickedFolder, tickedFile,checkboxLineEditFolderName, checkboxLineEditFileName):
        super().__init__()
        self.input_folder = input_folder
        self.output_folder = output_folder
        self.dpi_size = dpi_size
        self.num_images = num_images
        self.tickedFolder = tickedFolder
        self.tickedFile = tickedFile
        self.checkboxLineEditFolderName = checkboxLineEditFolderName
        self.checkboxLineEditFileName = checkboxLineEditFileName
    def run(self):
        processed_images = 0
        try:             # Convert and save the image as PNG
            for folder_name in os.listdir(self.input_folder):
                folder_path = os.path.join(self.input_folder, folder_name)
                folder_graph_path = os.path.join(folder_path, 'Graphics')  # Corrected path creation 'EMC32'
                if os.path.isdir(folder_graph_path):
                    for filename in os.listdir(folder_graph_path):
                        for i in range(len(self.tickedFolder)):
                            foldername = self.tickedFolder[i]
                            if folder_name.lower() == foldername.lower():
                                for j in range(len(self.tickedFile)):
                                    name = self.tickedFile[j]
                                    if filename.lower() == name.lower() + '.wmf':
                                        foldername = self.checkboxLineEditFolderName[i]
                                        name = self.checkboxLineEditFileName[j]
                                        with open(os.path.join(folder_graph_path, filename), "rb") as input_file:
                                            wmf_data = input_file.read()
                                        with WandImage(blob=wmf_data, resolution = int(self.dpi_size)) as img:
                                            foldername = str(foldername).lstrip('0')
                                            savename = (f"{self.output_folder}/{foldername[:4].replace('.','_',1) + foldername[4:]}_{name}.png")
                                            img.format = 'PNG'
                                            img.save(filename = savename)
                                            savename=''
                                        processed_images += 1
                                        self.progress = int((processed_images / self.num_images) * 100)
                                        self.progressChanged.emit(self.progress, processed_images, self.num_images)
            self.finished.emit()
            return         
        except Exception as e:
            raise Exception(f"Error extracting image '{self.input_folder}': {str(e)}")




class HighDpiFix:
    def __init__(self):
        if sys.platform == 'win32':
            if hasattr(Qt, 'AA_EnableHighDpiScaling'):
                QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)

            # Enable DPI awareness on Windows Vista (6.0) and later
            if sys.getwindowsversion().major >= 6:
                try:
                    from ctypes import windll
                    windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
                except ImportError:
                    pass

if __name__ == "__main__":
    high_dpi_fix = HighDpiFix()
    app = QApplication([])
    app.setFont(QFont('Serif', 10, QFont.Weight.Light))
    EMC32Extract = EMC32Extractor()
    app.exec()
