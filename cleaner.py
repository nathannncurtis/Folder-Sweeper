import sys
import os
import shutil
import xlrd
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,
                             QWidget, QLineEdit, QProgressBar, QMessageBox,
                             QFileDialog, QMenuBar, QAction)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon

class FolderProcessor(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, xls_dir, parent_dir, completed_dir, debug_mode=False):
        super().__init__()
        self.xls_dir = xls_dir
        self.parent_dir = parent_dir
        self.completed_dir = completed_dir
        self.debug_mode = debug_mode

    def debug_print(self, message):
        if self.debug_mode:
            print(message)

    def move_folder(self, source_folder, destination_folder):
        self.debug_print(f"Moving folder from {source_folder} to {destination_folder}")
        
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
            self.debug_print(f"Created destination folder: {destination_folder}")

        base_name = os.path.basename(source_folder)
        new_folder_path = os.path.join(destination_folder, base_name)
        counter = 1
        while os.path.exists(new_folder_path):
            self.debug_print(f"Folder {new_folder_path} already exists, renaming...")
            new_folder_path = os.path.join(destination_folder, f"{base_name} ({counter})")
            counter += 1

        shutil.move(source_folder, new_folder_path)
        self.debug_print(f"Folder moved successfully to {new_folder_path}")

    def run(self):
        self.debug_print("Thread started")
        self.debug_print(f"Started processing Excel files from {self.xls_dir}")
        
        folder_map = {folder_name: os.path.join(self.parent_dir, folder_name) 
                      for folder_name in os.listdir(self.parent_dir) 
                      if os.path.isdir(os.path.join(self.parent_dir, folder_name))}

        self.debug_print(f"Folder map created with {len(folder_map)} folders from parent directory")

        total_files = len([f for f in os.listdir(self.xls_dir) if f.endswith(('.xls', '.xlsx'))])
        self.debug_print(f"Total Excel files to process: {total_files}")
        progress_step = 100 / total_files if total_files > 0 else 100

        for i, file_name in enumerate(os.listdir(self.xls_dir)):
            if file_name.endswith(('.xls', '.xlsx')):
                file_path = os.path.join(self.xls_dir, file_name)
                self.debug_print(f"Processing file: {file_name}")
                
                try:
                    df = pd.read_excel(file_path, usecols=[0])
                    first_column_values = df.iloc[:, 0].dropna().astype(str).tolist()
                    self.debug_print(f"Extracted values from first column: {first_column_values}")

                    for value in first_column_values:
                        for folder_name, folder_path in folder_map.items():
                            if value in folder_name:
                                self.debug_print(f"Match found: {value} in folder {folder_name}")
                                modified_date = datetime.fromtimestamp(os.path.getmtime(folder_path))
                                month_folder = modified_date.strftime("%m %B %Y")
                                completed_subfolder = os.path.join(self.completed_dir, month_folder)
                                self.debug_print(f"Moving folder {folder_name} to completed subfolder {completed_subfolder}")
                                self.move_folder(folder_path, completed_subfolder)

                except Exception as e:
                    self.debug_print(f"Error processing {file_name}: {str(e)}")

            self.progress.emit(int((i + 1) * progress_step))

        self.debug_print("Processing finished")
        self.finished.emit()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.debug_mode = False  # debug mode off by default, not persistent between sessions either
        self.initUI()
        self.loadSettings()

    def initUI(self):
        self.setWindowTitle("Cleaner")
        self.setFixedSize(400, 300)
        self.setWindowIcon(QIcon('cleaner.ico'))

        layout = QVBoxLayout()

        # default path
        default_parent_path = r'\\ronsyn\ClientServices'
        default_completed_path = r'\\ronsyn\ClientServices\_COMPLETED'

        self.xls_input = QLineEdit()
        self.xls_input.setPlaceholderText("Select directory with .xls/.xlsx files...")
        layout.addWidget(self.xls_input)

        self.browse_xls_button = QPushButton("Browse for Invoiced Orders (.xls/.xlsx)")
        self.browse_xls_button.clicked.connect(self.browse_xls)
        layout.addWidget(self.browse_xls_button)

        self.parent_input = QLineEdit(default_parent_path)
        self.parent_input.setPlaceholderText("Select the Folder to Clean")
        layout.addWidget(self.parent_input)

        self.browse_parent_button = QPushButton("Browse for Parent Folder")
        self.browse_parent_button.clicked.connect(self.browse_parent)
        layout.addWidget(self.browse_parent_button)

        self.completed_input = QLineEdit(default_completed_path)
        self.completed_input.setPlaceholderText("Select Completed Folder...")
        layout.addWidget(self.completed_input)

        self.browse_completed_button = QPushButton("Browse for Completed Folder")
        self.browse_completed_button.clicked.connect(self.browse_completed)
        layout.addWidget(self.browse_completed_button)

        self.start_button = QPushButton("Start Cleaning")
        self.start_button.clicked.connect(self.start_processing)
        layout.addWidget(self.start_button)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.progress_bar)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.menuBar = self.menuBar()
        extrasMenu = self.menuBar.addMenu('Extras')

        self.toggleThemeAction = QAction('Toggle Theme', self)
        self.toggleThemeAction.triggered.connect(self.toggle_theme)
        extrasMenu.addAction(self.toggleThemeAction)

        self.debugAction = QAction('Enable Debug Mode', self)
        self.debugAction.setCheckable(True)
        self.debugAction.triggered.connect(self.toggle_debug_mode)
        extrasMenu.addAction(self.debugAction)

    def loadSettings(self):
        self.dark_mode = True
        self.apply_theme()

    def toggle_theme(self):
        self.debug_print("Toggling theme")
        self.dark_mode = not self.dark_mode
        self.apply_theme()

    def apply_theme(self):
        self.debug_print(f"Applying {'dark' if self.dark_mode else 'light'} theme")
        if self.dark_mode:
            self.setStyleSheet("""
                QMainWindow, QDialog {
                    background-color: #333;
                    color: white;
                }
                QLineEdit, QPushButton, QProgressBar, QMenuBar, QMenu {
                    background-color: #555;
                    color: white;
                }
                QLabel {
                    color: white;
                }
                QProgressBar {
                    border: 1px solid #666;
                    background-color: #333;
                }
                QProgressBar::chunk {
                    background-color: #06b;
                }
                QMenuBar::item:selected {
                    background-color: #06b;
                }
                QMenu::item:selected {
                    background-color: #333;
                }
            """)
        else:
            self.setStyleSheet("""
                QMainWindow, QDialog {
                    background-color: #eee;
                    color: black;
                }
                QLineEdit, QPushButton, QProgressBar, QMenuBar, QMenu {
                    background-color: #ccc;
                    color: black;
                }
                QLabel {
                    color: black;
                }
                QProgressBar {
                    border: 1px solid #bbb;
                    background-color: #eee;
                }
                QProgressBar::chunk {
                    background-color: #06b;
                }
                QMenuBar::item:selected {
                    background-color: #a0c4ff;
                }
                QMenu::item:selected {
                    background-color: #a0c4ff;
                }
            """)

    def browse_xls(self):
        self.debug_print("Browsing for Excel files directory")
        directory = QFileDialog.getExistingDirectory(self, "Select Folder with Invoice Reports")
        if directory:
            self.xls_input.setText(directory)

    def browse_parent(self):
        self.debug_print("Browsing for parent folder")
        directory = QFileDialog.getExistingDirectory(self, "Select Parent Folder with Files to Clean")
        if directory:
            self.parent_input.setText(directory)

    def browse_completed(self):
        self.debug_print("Browsing for completed folder")
        directory = QFileDialog.getExistingDirectory(self, "Select Completed Folder")
        if directory:
            self.completed_input.setText(directory)

    def start_processing(self):
        xls_dir = self.xls_input.text()
        parent_dir = self.parent_input.text()
        completed_dir = self.completed_input.text()

        if not (xls_dir and parent_dir and completed_dir):
            self.debug_print("Error: Not all directory paths are filled in")
            QMessageBox.warning(self, "Error", "Please fill in all directory paths.", QMessageBox.Ok)
            return

        self.debug_print("Starting processing thread")
        self.progress_bar.setValue(0)
        self.processor = FolderProcessor(xls_dir, parent_dir, completed_dir, self.debug_mode)
        self.processor.progress.connect(self.progress_bar.setValue)
        self.processor.finished.connect(self.processing_finished)
        self.processor.start()

    def processing_finished(self):
        self.debug_print("Processing finished")
        QMessageBox.information(self, "Finished", "The folder is squeaky clean.", QMessageBox.Ok)

    def toggle_debug_mode(self):
        self.debug_mode = not self.debug_mode
        self.debug_print(f"Debug mode {'enabled' if self.debug_mode else 'disabled'}")
        if self.debug_mode:
            self.debugAction.setText("Disable Debug Mode")
            self.show_console()
        else:
            self.debugAction.setText("Enable Debug Mode")
            self.hide_console()

    def show_console(self):
        # allocates a console for debug prints — windows only
        if sys.platform == 'win32':
            import ctypes
            ctypes.windll.kernel32.AllocConsole()
            self.debug_print("Console allocated for debug prints")

    def hide_console(self):
        # close console — windows only 
        if sys.platform == 'win32':
            import ctypes
            ctypes.windll.kernel32.FreeConsole()
            self.debug_print("Console closed")

    def debug_print(self, message):
        if self.debug_mode:
            print(message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.exit(app.exec_())