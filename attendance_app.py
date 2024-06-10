import sys
import os
import pandas as pd
from datetime import datetime, timedelta
from dateutil.parser import parse
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox, QComboBox

def convert_xls_to_xlsx(input_path, output_path):
    df = pd.read_excel(input_path)
    df.to_excel(output_path, index=False)

def preprocess_data(df, no_column, time_column):
    df.columns = df.columns.str.lower()
    no_column = no_column.lower()
    time_column = time_column.lower()
    if no_column not in df.columns or time_column not in df.columns:
        raise ValueError("Input file must contain the selected columns for employee number and attendance date/time.")
    df = df.rename(columns={no_column: 'Employee', time_column: 'Check in'})
    df['Check in'] = df['Check in'].str.replace('ص', 'AM').str.replace('م', 'PM').str.replace('Õ', 'AM').str.replace('ã', 'PM')
    df['Check in'] = df['Check in'].apply(lambda x: parse(x, dayfirst=True, fuzzy=True) if pd.notnull(x) else x)
    df = df.sort_values(by=['Employee', 'Check in'])
    return df

def process_attendance(df):
    processed_records = []
    df['Date'] = df['Check in'].dt.date
    grouped = df.groupby(['Employee', 'Date'])
    for (employee, date), group in grouped:
        group = group.reset_index(drop=True)
        shift_start = None
        shift_end = None
        for i in range(len(group)):
            current_time = group.loc[i, 'Check in']
            if shift_start is None:
                shift_start = current_time
                shift_end = None
                continue
            if current_time - shift_start < timedelta(minutes=30):
                continue
            if shift_end is None or current_time - shift_end >= timedelta(hours=1):
                if shift_end is not None:
                    processed_records.append((employee, shift_start, shift_end))
                    shift_start = current_time
                    shift_end = None
                else:
                    shift_end = current_time
        if shift_end is not None:
            processed_records.append((employee, shift_start, shift_end))
        else:
            processed_records.append((employee, shift_start, shift_start))
    processed_df = pd.DataFrame(processed_records, columns=['Employee', 'Check in', 'Check out'])
    return processed_df

class AttendanceApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel("Upload Attendance File (xls, xlsx, csv)")
        layout.addWidget(self.label)

        self.upload_btn = QPushButton("Upload File")
        self.upload_btn.clicked.connect(self.upload_file)
        layout.addWidget(self.upload_btn)

        self.no_label = QLabel("Select Employee Number Column")
        layout.addWidget(self.no_label)
        self.no_combo = QComboBox()
        layout.addWidget(self.no_combo)

        self.time_label = QLabel("Select Time Column")
        layout.addWidget(self.time_label)
        self.time_combo = QComboBox()
        layout.addWidget(self.time_combo)

        self.process_btn = QPushButton("Process File")
        self.process_btn.clicked.connect(self.process_file)
        layout.addWidget(self.process_btn)

        self.setLayout(layout)
        self.setWindowTitle('Attendance Processor')
        self.setGeometry(300, 300, 400, 300)

    def upload_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "", "All Files (*);;Excel Files (*.xls *.xlsx);;CSV Files (*.csv)", options=options)
        if file_name:
            self.file_path = file_name
            self.label.setText(f"File Selected: {file_name}")
            if file_name.endswith('.xls'):
                output_path = file_name.replace('.xls', '.xlsx')
                convert_xls_to_xlsx(file_name, output_path)
                file_name = output_path
            self.df = pd.read_excel(file_name) if file_name.endswith('.xlsx') else pd.read_csv(file_name)
            self.no_combo.clear()
            self.no_combo.addItems(self.df.columns)
            self.time_combo.clear()
            self.time_combo.addItems(self.df.columns)

    def process_file(self):
        if hasattr(self, 'df'):
            try:
                no_column = self.no_combo.currentText()
                time_column = self.time_combo.currentText()
                df = preprocess_data(self.df, no_column, time_column)
                processed_df = process_attendance(df)
                base_name, _ = os.path.splitext(self.file_path)
                current_date_time = datetime.now().strftime('%Y%m%d_%H%M%S')
                output_filename = f"{base_name}_{current_date_time}.xlsx"
                processed_df.to_excel(output_filename, index=False)
                QMessageBox.information(self, "Success", f"Processed file saved as {output_filename}")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
        else:
            QMessageBox.warning(self, "Warning", "Please upload a file first")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = AttendanceApp()
    ex.show()
    sys.exit(app.exec_())
