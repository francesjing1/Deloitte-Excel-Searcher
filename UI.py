from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QListWidget, QPushButton, QLineEdit, QVBoxLayout,
    QFileDialog, QMessageBox
)
import sys

from Searcher import Searcher

class UI(QWidget):
    def __init__(self, process_files_callback):
        super().__init__()
        self.setWindowTitle("Searcher")
        self.setGeometry(300, 300, 800, 550)

        self.files = []
        self.process_files_callback = process_files_callback

        # Widgets
        self.label = QLabel("Select Excel files (.xlsb only):")
        self.listbox = QListWidget()
        self.add_button = QPushButton("Add Files")
        self.keyword_label = QLabel("Enter keyword:")
        self.keyword_entry = QLineEdit()
        self.process_button = QPushButton("Search Files")

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.listbox)
        layout.addWidget(self.add_button)
        layout.addWidget(self.keyword_label)
        layout.addWidget(self.keyword_entry)
        layout.addWidget(self.process_button)
        self.setLayout(layout)

        # Connections
        self.add_button.clicked.connect(self.add_files)
        self.process_button.clicked.connect(self.process_files)

    def add_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Excel Files",
            "",
            "Excel files (*.xlsb)"
        )
        for file_path in file_paths:
            if file_path not in self.files:
                self.files.append(file_path)
                self.listbox.addItem(file_path)

    def process_files(self):
        keyword = self.keyword_entry.text().strip()
        if not self.files:
            QMessageBox.warning(self, "No Files", "Please add at least one Excel file.")
            return
        if not keyword:
            QMessageBox.warning(self, "No Keyword", "Please enter a keyword.")
            return
        self.process_files_callback(self.files, keyword)

def process_selected_files(file_list, keyword):
    for file in file_list:
        Searcher.readExcelFile(file, keyword)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ui = UI(process_selected_files)
    ui.show()
    sys.exit(app.exec_())

