import os
import re
import csv
import fitz  # PyMuPDF
from datetime import datetime
import sys
import pandas as pd
from PyQt6.QtWidgets import QApplication, QWidget, QComboBox, QLineEdit, QTableView, QPushButton, QLabel, \
							QHBoxLayout, QVBoxLayout
from PyQt6.QtCore import Qt, QAbstractTableModel, QVariant, QModelIndex
from PyQt6.QtGui import QIcon
               
# Define the static text to ignore
IGNORE_TEXT = {
    "DUTY DESCRIPTION": [
        "RATER ASSESSMENT"
    ],
    "EXECUTING THE MISSION": [
        "EFFECTIVELY USES KNOWLEDGE, INITIATIVE, AND ADAPTABILITY TO PRODUCE TIMELY, HIGH QUALITY/QUANTITY RESULTS TO POSITIVELY IMPACT THE MISSION"
    ],
    "LEADING PEOPLE": [
        "FOSTERS COHESIVE TEAMS, EFFECTIVELY COMMUNICATES, AND USES EMOTIONAL INTELLIGENCE TO TAKE CARE OF PEOPLE AND ACCOMPLISH THE MISSION"
    ],
    "MANAGING RESOURCES": [
        "MANAGES ASSIGNED RESOURCES EFFECTIVELY AND TAKES RESPONSIBILITY FOR ACTIONS/BEHAVIORS TO MAXIMIZE ORGANIZATIONAL PERFORMANCE"
    ],
    "IMPROVING THE UNIT": [
        "DEMONSTRATES CRITICAL THINKING AND FOSTERS INNOVATION TO FIND CREATIVE SOLUTIONS AND IMPROVE MISSION EXECUTION"
    ],
    "HIGHER LEVEL REVIEWER ASSESSMENT": [
        "RATER SIGNATURE"
    ]
}

CATEGORIES = [
    "DUTY DESCRIPTION",
    "EXECUTING THE MISSION",
    "LEADING PEOPLE",
    "MANAGING RESOURCES",
    "IMPROVING THE UNIT",
    "HIGHER LEVEL REVIEWER ASSESSMENT"
]

def parse_pdf(file_path):
    file_name = os.path.basename(file_path).split('-')[0]  # Extract name before hyphen
    
    document = fitz.open(file_path)
    text = ""
    for page_num in range(document.page_count):
        page = document.load_page(page_num)
        text += page.get_text("text")

    lines = text.splitlines()
    categorized_statements = []
    current_category = None
    statement = []
    recent_lines = []  # Store last 3 lines for HIGHER LEVEL REVIEWER ASSESSMENT
    days_supervised = None  # To store the number of days supervised
    duty_title = None  # To store the duty title
    duty_title_found = False  # Flag to check if we have already found the first "DUTY TITLE"
    dafsc = None  # To store the DAFSC
    dafsc_found = False  # Flag to check if we have already found the first "DAFSC"
    reason = None  # To store the reason
    reason_found = False  # Flag to check if we have already found the first "REASON"
    org = None  # To store the organization
    org_found = False  # Flag to check if we have already found the first "ORGANIZATION AND COMMAND"
    period_start = None  # To store the period start date
    period_end = None  # To store the period end date
    period_found = False  # Flag to check if we have already found the first "PERIOD"
    location = None  # To store the location
    location_found = False  # Flag to check if we have already found the first "LOCATION"
    ratee_signed = None  # To store the ratee signed date
    ratee_signed_found = False  # Flag to check if we have already found the first "RATEE ACKNOWLEDGEMENT"
    rater_name = None  # To store the rater name
    rater_name_found = False  # Flag to check if we have already found the first "RATER NAME, GRADE, AND BRANCH OF SERVICE"
    rater_signed = None  # To store the rater signed date
    rater_signed_found = False  # Flag to check if we have already found the first "RATER SIGNATURE"
    rater_duty_title = None  # To store the RATER DUTY TITLE
    rater_duty_title_found = False  # Flag to check if we have already found the first "DUTY TITLE"
    hlr_name = None  # To store the HIGHER LEVEL REVIEWER NAME
    hlr_name_found = False  # Flag to check if we have already found the first "HIGHER LEVEL REVIEWER NAME, GRADE, AND BRANCH OF SERVICE"
    hlr_duty_title_found = False  # Flag to check if we have already found the first "HIGHER LEVEL REVIEWER DUTY TITLE"
    hlr_duty_title = None  # To store the HIGHER LEVEL REVIEWER DUTY TITLE
    hlr_signed = None  # To store the HIGHER LEVEL REVIEWER SIGNATURE
    hlr_signed_found = False  # Flag to check if we have already found the first "HIGHER LEVEL REVIEWER SIGNATURE"
    strat = None  # To store the STRATIFICATION
    strat_found = False  # Flag to check if we have already found the first "STRATIFICATION"
    future_roles = {"1": None, "2": None, "3": None}  # Store future roles
    promotion_recommendation = None  # To store the promotion recommendation
    promotion_recommendation_found = False  # Flag to check if we have already found the first "PROMOTION RECOMMENDATION"
    days_non_rated = None  # To store the number of days non-rated

    for i, line in enumerate(lines):
        
       
        for category in CATEGORIES:
            if line.startswith(category):
                if current_category and statement:
                    filtered_statement = remove_unwanted_text(current_category, ' '.join(statement))
                    categorized_statements.extend(split_sentences(current_category, filtered_statement, file_path, file_name))
                
                current_category = category
                statement = [line[len(category):].strip()]
                recent_lines = []  # Reset recent lines
                break
        else:
            if any(ignored_text in line for ignored_text in IGNORE_TEXT.get(current_category, [])):
                continue
            
            if current_category == "HIGHER LEVEL REVIEWER ASSESSMENT":
                if "HIGHER LEVEL REVIEWER NAME, GRADE, AND BRANCH OF SERVICE" in line:
                    # Capture only the last 3 lines before this marker
                    statement = recent_lines[-3:]  # Get the last 3 non-empty lines
                    filtered_statement = remove_unwanted_text(current_category, ' '.join(statement))
                    # Split and add this statement as individual records
                    categorized_statements.extend(split_sentences(current_category, filtered_statement, file_path, file_name))
                    current_category = None  # Stop further processing for this category
                
                
                elif "FUTURE ROLES" in line:
                    future_roles["1"] = lines[i + 1].strip().split("1.", 1)[1].strip()  # Extract role 1
                    future_roles["2"] = lines[i + 2].strip().split("2.", 1)[1].strip()  # Extract role 2
                    future_roles["3"] = lines[i + 3].strip().split("3.", 1)[1].strip()  # Extract role 3
                else:
                    recent_lines.append(line.strip())  
                    recent_lines = recent_lines[-5:]  # Keep only the last 5 lines (extra buffer)

            elif current_category:
                statement.append(line.strip())
            
            if "HIGHER LEVEL REVIEWER DUTY TITLE" in line and not hlr_duty_title_found and i + 1 < len(lines):
                    # Capture the line following "HIGHER LEVEL REVIEWER DUTY TITLE"
                    hlr_duty_title = lines[i + 1].strip()
                    hlr_duty_title_found = True
            
            if "HIGHER LEVEL REVIEWER NAME, GRADE, AND BRANCH OF SERVICE" in line and not hlr_name_found and i + 1 < len(lines):
                hlr_name = lines[i + 1].strip()
                hlr_name_found = True
            
            if "RATER NAME, GRADE, AND BRANCH OF SERVICE" in line and not rater_name_found and i + 1 < len(lines):
                rater_name = lines[i + 1].strip()
                rater_name_found = True
            
            if "HIGHER LEVEL REVIEWER SIGNATURE" in line and not hlr_signed_found and i + 1 < len(lines):
                if "HIGHER LEVEL REVIEWER DUTY TITLE" in lines[i + 1]:
                    hlr_signature = ""
                else:
                    hlr_signature = lines[i + 1].strip()
                    parts = hlr_signature.split(',')
                
                if len(parts) > 3:
                    hlr_signed = parts[3].strip().strip('\\')  # Extract and clean the date

                    # Validate format
                    if not is_valid_date(hlr_signed):
                        print(f"Invalid higher-level reviewer signed date format: {hlr_signed}")
                        hlr_signed = None  # Handle invalid date as needed

                hlr_signed_found = True

            if "RATER SIGNATURE" in line and not rater_signed_found and i + 1 < len(lines):
                rater_signature = lines[i + 1].strip()
                parts = rater_signature.split(',')
                
                if len(parts) > 3:
                    rater_signed = parts[3].strip().strip('\\')  # Extract and clean the date

                    # Validate format
                    if not is_valid_date(rater_signed):
                        print(f"Invalid rater signed date format: {rater_signed}")
                        rater_signed = None  # Handle invalid date as needed

                rater_signed_found = True

            if "RATER DUTY TITLE" in line and not rater_duty_title_found and i + 1 < len(lines):
                rater_duty_title = lines[i + 1].strip()  # Capture the line following "RATER DUTY TITLE"
                rater_duty_title_found = True  # Set flag to indicate we've found the first instance of "RATER DUTY TITLE"

            if "RATEE ACKNOWLEDGEMENT" in line and not ratee_signed_found and i + 1 < len(lines):
                if "ORGANIZATION AND COMMAND" in lines[i + 1]:
                    ratee_signature = ""
                else:
                    ratee_signature = lines[i + 1].strip()
                    parts = ratee_signature.split(',')
                
                if len(parts) > 3:
                    ratee_signed = parts[3].strip().strip('\\')  # Extract the date and clean up backslashes

                    # Validate format
                    if not is_valid_date(ratee_signed):
                        print(f"Invalid ratee signed date format: {ratee_signed}")
                        ratee_signed = None  # Set to None or handle as needed

                ratee_signed_found = True  # Set flag to indicate we've found the first instance of "RATEE ACKNOWLEDGEMENT"
            if "STRATIFICATION" in line and not strat_found and i + 1 < len(lines):
                    if "FORCED ENDORSEMENT" in lines[i + 1]:
                        strat = ""
                    else:
                        strat = lines[i + 1].strip()  # Capture the line following "STRATIFICATION"
                    strat_found = True
            if "PROMOTION RECOMMENDATION" in line and not promotion_recommendation_found and i + 1 < len(lines):
                if "RATER ASSESSMENT" in lines[i + 1]:
                    promotion_recommendation = ""
                else:
                    promotion_recommendation = lines[i + 1].strip()
                promotion_recommendation_found = True

            # Look for the "DAYS SUPERVISED" text and extract the next three numbers
            if "DAYS SUPERVISED" in line and i + 1 < len(lines):
                days_supervised = extract_number_from_text(lines[i + 1]) 
            
            if "DAYS NON-RATED" in line and i + 1 < len(lines):
                days_non_rated = extract_number_from_text(lines[i + 1])

            # Look for the first "DUTY TITLE" text and extract the next line
            if "DUTY TITLE" in line and not duty_title_found and i + 1 < len(lines):
                if "DAFSC" in lines[i + 1]:
                    duty_title = ""
                else:
                    duty_title = lines[i + 1].strip()  # Capture the line following "DUTY TITLE"
                duty_title_found = True  # Set flag to indicate we've found the first instance of "DUTY TITLE"

            # Look for the first "DAFSC" text and extract the next line
            if "DAFSC" in line and not dafsc_found and i + 1 < len(lines):
                dafsc = lines[i + 1].strip()  # Capture the line following "DAFSC"
                dafsc_found = True  # Set flag to indicate we've found the first instance of "DAFSC"

            # Look for the first "REASON" text and extract the next line
            if "REASON" in line and not reason_found and i + 1 < len(lines):
                reason = lines[i + 1].strip()  # Capture the line following "REASON"
                reason_found = True  # Set flag to indicate we've found the first instance of "REASON"

            # Look for the first "ORGANIZATION AND COMMAND" text and extract the next line
            if "ORGANIZATION AND COMMAND" in line and not org_found and i + 1 < len(lines):
                org = lines[i + 1].strip()  # Capture the line following "ORGANIZATION AND COMMAND"
                org_found = True 

            # Look for the first "LOCATION" text and extract the next line
            if "LOCATION" in line and not location_found and i + 1 < len(lines):
                location = lines[i + 1].strip()  # Capture the line following "LOCATION"
                location_found = True

            if "PERIOD" in line and not period_found and i + 1 < len(lines):
                period_text = lines[i + 1].strip()  # Capture the line following "PERIOD"
                if "THRU" in period_text:
                    period_start, period_end = period_text.split("THRU", 1)
                    period_start = period_start.strip()
                    period_end = period_end.strip()

                    # Validate format
                    if not is_valid_date(period_start):
                        period_start = None  # Set to None or handle as needed
                    if not is_valid_date(period_end):
                        period_end = None  # Set to None or handle as needed

                period_found = True  # Set flag to indicate we've found the first instance of "PERIOD" 

    if current_category and statement:
        filtered_statement = remove_unwanted_text(current_category, ' '.join(statement))
        categorized_statements.extend(split_sentences(current_category, filtered_statement, file_path, file_name))
       
    # Add the extracted values to each record for this PDF
    for record in categorized_statements:
      
        record.append(days_supervised)  # Append days_supervised to each record
        record.append(days_non_rated)  # Append days_non_rated to each record
        record.append(duty_title)  # Append duty_title to each record
        record.append(dafsc)  # Append dafsc to each record
        record.append(reason)  # Append reason to each record
        record.append(period_start)  # Append period_start to each record
        record.append(period_end)  # Append period_end to each record
        record.append(org)  # Append org to each record
        record.append(location)  # Append location to each record
        record.append(ratee_signed)  # Append ratee_signed to each record
        record.append(rater_name)  # Append rater_name to each record
        record.append(rater_signed)  # Append rater_signed to each record
        record.append(rater_duty_title)  # Append rater_duty_title to each record
        record.append(hlr_name)  # Append hlr_name to each record
        record.append(hlr_duty_title)  # Append hlr_duty_title to each record
        record.append(hlr_signed)  # Append hlr_signed to each record
        record.append(strat)  # Append strat to each record
        record.append(promotion_recommendation)  # Append promotion_recommendation to each record
        record.append(future_roles["1"])  # Append future_role_1 to each record
        record.append(future_roles["2"])  # Append future_role_2 to each record
        record.append(future_roles["3"])  # Append future_role_3 to each record
        
    return categorized_statements

def extract_number_from_text(text):
    """
    Extracts the first number found in the given text, if any.
    """
    match = re.findall(r'\d+', text)
    if match:
        return match[0]  # Return the first number found as a string
    return None

def split_sentences(category, statement, file_path, file_name):
    """
    Splits a statement into multiple records if it contains more than one complete sentence.
    Ensures proper handling of numeric values with periods.
    """
    # If the category is "HIGHER LEVEL REVIEWER ASSESSMENT", return the statement as is
    if category == "HIGHER LEVEL REVIEWER ASSESSMENT":
        return [[category, statement, file_path, file_name]]
    
    # For other categories, proceed with splitting the statement into sentences

    # First, replace common abbreviations and numeric periods with a placeholder to prevent splitting
    statement = re.sub(r'\b(?:Dr|Mr|Ms|U\.S\.)\b', lambda match: match.group(0).replace('.', '__DOT__'), statement)
    statement = re.sub(r'(\d)(?=\.\d)', r'\1__DOT__', statement)  # Handle numeric periods like 1.4

    # Now split the statement on standard sentence-ending punctuation
    sentences = re.split(r'(?<=[.!?])\s+', statement)

    # Restore the placeholders to their original form
    sentences = [s.replace('__DOT__', '.') for s in sentences]

    processed_statements = []
    for sentence in sentences:
        sentence = sentence.strip()
        if sentence and not sentence.endswith(('.', '!', '?')):
            sentence += '.'  # Ensure it ends with a period
        processed_statements.append([category, sentence, file_path, file_name])

    return processed_statements

def is_valid_date(date_str):
    """Check if a date matches the format 'D Mmm YY' or 'DD Mmm YY' and is a real date."""
    pattern = r"^\d{1,2} [A-Za-z]{3} \d{2}$"  # Allows 1 or 2 digits for the day
    if not re.match(pattern, date_str):
        return False
    try:
        datetime.strptime(date_str, "%d %b %y")  # Validates real date
        return True
    except ValueError:
        return False

def remove_unwanted_text(category, statement):
    unwanted_text = IGNORE_TEXT.get(category, [])
    
    for text in unwanted_text:
        statement = statement.replace(text, "")

    if category == "IMPROVING THE UNIT":
        truncation_point = "RATER NAME, GRADE, AND BRANCH OF SERVICE"
        if truncation_point in statement:
            statement = statement.split(truncation_point)[0]  

    statement = re.sub(r'\s+', ' ', statement).strip()  
    return statement

def save_to_csv(data, output_file):
    # Ensure the header includes the additional fields
    with open(output_file, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([
            "Category", "Statement", "file_path", "Name", "days_supervised", "days_non_rated", "duty_title", 
            "dafsc", "reason", "period_start", "period_end", "org", "location", 
            "ratee_signed", "rater_name", "rater_signed", "rater_duty_title", "HLR_name", 
            "HLR_duty_title", "HLR_signed", "strat","promotion_rec", "future_role_1", "future_role_2", "future_role_3"
        ])  # Add all headers here
        writer.writerows(data)

# Main script


class PandasModel(QAbstractTableModel):
	def __init__(self, df=pd.DataFrame(), parent=None):
		super().__init__(parent)
		self._df = df

	def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
		if role != Qt.ItemDataRole.DisplayRole:
			return QVariant()
		if orientation == Qt.Orientation.Horizontal:
			try:
				return self._df.columns.tolist()[section]
			except IndexError:
				return QVariant()
		elif orientation == Qt.Orientation.Vertical:
			try:
				return self._df.index.tolist()[section]
			except IndexError:
				return QVariant()

	def rowCount(self, parent=QModelIndex()):
		return self._df.shape[0]

	def columnCount(self, parent=QModelIndex()):
		return self._df.shape[1]

	def data(self, index, role=Qt.ItemDataRole.DisplayRole):
		if role != Qt.ItemDataRole.DisplayRole:
			return QVariant()
		if not index.isValid():
			return QVariant()
		return QVariant(str(self._df.iloc[index.row(), index.column()]))



class MyApp(QWidget):

    def __init__(self, defaultSource=None):
        super().__init__()
        self.window_width, self.window_height = 1100, 500
        self.resize(self.window_width, self.window_height)
        self.setWindowTitle('CSV Data Viewer')
        self.setWindowIcon(QIcon('./icon/browser.png'))
        self.df = None
        self.setStyleSheet("""
            QWidget {
                font-size: 15px;
            }
            QComboBox {
                width: 160px;
            }
            QPushButton {
                width: 100px;
            }
        """)

        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.initUI(defaultSource)

    def retrieveDataset(self):
        try:
            urlSource = self.dataSourceField.text()
            self.df = pd.read_csv(urlSource)
            self.df.fillna('')
            self.model = PandasModel(self.df)
            self.table.setModel(self.model)

            self.comboColumns.clear()
            self.comboColumns.addItems(self.df.columns)
        except Exception as e:
            self.statusLabel.setText(str(e))
            return

    def searchItem(self, v):
        if self.df is None:
            return

        column_index = self.df.columns.get_loc(self.comboColumns.currentText())
        for row_index in range(self.model.rowCount()):
            if v in self.model.index(row_index, column_index).data():
                self.table.setRowHidden(row_index, False)
            else:
                self.table.setRowHidden(row_index, True)

    def copy_to_clipboard(self):
        try:
            clipboard = QApplication.clipboard()
            visible_rows = []

            # Iterate over the rows and check if they are visible
            for row_index in range(self.model.rowCount()):
                if not self.table.isRowHidden(row_index):
                    # Extract only the value of the "Statements" column (index 1)
                    statement = str(self.model.data(self.model.index(row_index, 1)).value())  # Column 1 is for Statements
                    visible_rows.append([statement])

            if not visible_rows:
                print("No visible rows to copy.")
                return

            # Create CSV string for the Statements column
            csv_data = "\n".join([row[0] for row in visible_rows])  # Only join the statements, not the whole row

            # Debug output to ensure data is correct
            print("Statements to copy:", visible_rows)
            
            # Set the clipboard data
            clipboard.setText(csv_data)
            print("Statements copied to clipboard successfully.")

        except Exception as e:
            print(f"An error occurred while copying to clipboard: {e}")
            

    def initUI(self, defaultSource):
        sourceLayout = QHBoxLayout()
        self.layout.addLayout(sourceLayout)

        label = QLabel('&Data Source: ')
        self.dataSourceField = QLineEdit(defaultSource)  # Default value set here
        label.setBuddy(self.dataSourceField)

        buttonRetrieve = QPushButton('&Retrieve', clicked=self.retrieveDataset)
        buttonCopy = QPushButton('&Copy Statements', clicked=self.copy_to_clipboard)

        sourceLayout.addWidget(label)
        sourceLayout.addWidget(self.dataSourceField)
        sourceLayout.addWidget(buttonRetrieve)

        # search field
        
        searchLayout = QHBoxLayout()
        self.layout.addLayout(searchLayout)

        label = QLabel('&Search: ')
        self.searchField = QLineEdit()
        self.searchField.textChanged.connect(self.searchItem)
        label.setBuddy(self.searchField)
        searchLayout.addWidget(label)
        searchLayout.addWidget(self.searchField)
        searchLayout.addWidget(buttonCopy)

        self.comboColumns = QComboBox()
        
        searchLayout.addWidget(self.comboColumns)

        self.table = QTableView()
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionsMovable(True)
        self.layout.addWidget(self.table)

        self.statusLabel = QLabel()
        self.statusLabel.setText('')
        self.layout.addWidget(self.statusLabel)

        # Automatically call the retrieveDataset function after UI is initialized
        self.retrieveDataset()

if __name__ == '__main__':
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    output_csv = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "output.csv")
    

    all_statements = []
    for pdf_file in os.listdir(script_dir):
        if pdf_file.endswith(".pdf"):
            file_path = os.path.join(script_dir, pdf_file)
            statements = parse_pdf(file_path)
            all_statements.extend(statements)

    save_to_csv(all_statements, output_csv)

    app = QApplication(sys.argv)
    app.setStyleSheet('''
		QWidget {
			font-size: 17px;
		}
	''')
	
    myApp = MyApp(output_csv)
    myApp.show()

    try:
        sys.exit(app.exec())
    except SystemExit:
        print('Closing Window...')

print(f"Parsing complete. Results saved to {output_csv}")