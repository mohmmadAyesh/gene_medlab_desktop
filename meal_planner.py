import sys
import random
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QSpinBox, QPushButton, 
                           QTableWidget, QTableWidgetItem, QMessageBox,
                           QGroupBox, QFileDialog, QScrollArea, QCheckBox,
                           QTabWidget, QComboBox)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from meal_items import MEAL_ITEMS

class MealPlanner(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Meal Planner")
        self.setMinimumSize(1200, 800)
        
        # Initialize the main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # Create tab widget
        tabs = QTabWidget()
        
        # Create main tab
        main_tab = QWidget()
        main_layout = QVBoxLayout(main_tab)
        
        # Create input section
        input_group = QGroupBox("Category Settings")
        input_layout = QHBoxLayout()
        
        # Category A input
        category_a_layout = QVBoxLayout()
        category_a_label = QLabel("Category A Count:")
        self.category_a_spin = QSpinBox()
        self.category_a_spin.setRange(1, 6)
        self.category_a_spin.setValue(4)
        category_a_layout.addWidget(category_a_label)
        category_a_layout.addWidget(self.category_a_spin)
        
        # Category B input
        category_b_layout = QVBoxLayout()
        category_b_label = QLabel("Category B Count:")
        self.category_b_spin = QSpinBox()
        self.category_b_spin.setRange(1, 6)
        self.category_b_spin.setValue(2)
        category_b_layout.addWidget(category_b_label)
        category_b_layout.addWidget(self.category_b_spin)
        
        # Category C input
        category_c_layout = QVBoxLayout()
        category_c_label = QLabel("Category C Count:")
        self.category_c_spin = QSpinBox()
        self.category_c_spin.setRange(1, 6)
        self.category_c_spin.setValue(1)
        category_c_layout.addWidget(category_c_label)
        category_c_layout.addWidget(self.category_c_spin)
        
        input_layout.addLayout(category_a_layout)
        input_layout.addLayout(category_b_layout)
        input_layout.addLayout(category_c_layout)
        input_group.setLayout(input_layout)
        
        # Button layout
        button_layout = QHBoxLayout()
        self.generate_button = QPushButton("Generate Meal Plan")
        self.generate_button.clicked.connect(self.generate_meal_plan)
        self.save_button = QPushButton("Save to Excel")
        self.save_button.clicked.connect(self.save_to_excel)
        self.save_button.setEnabled(False)
        button_layout.addWidget(self.generate_button)
        button_layout.addWidget(self.save_button)
        
        # Table widget for displaying meal plan
        table_group = QGroupBox("Meal Plan")
        table_layout = QVBoxLayout()
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["اليوم", "الإفطار", "الغداء", "العشاء"])
        self.table.horizontalHeader().setStretchLastSection(True)
        table_layout.addWidget(self.table)
        table_group.setLayout(table_layout)
        
        # Add widgets to main tab layout
        main_layout.addWidget(input_group)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(table_group)
        
        # Create exclusion tab
        exclusion_tab = QWidget()
        exclusion_layout = QVBoxLayout(exclusion_tab)
        
        # Create scroll area for exclusions
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        # Create checkboxes for each meal item
        self.exclusion_checkboxes = {}
        for item in MEAL_ITEMS:
            checkbox = QCheckBox(item["name"])
            checkbox.setChecked(False)
            self.exclusion_checkboxes[item["name"]] = checkbox
            scroll_layout.addWidget(checkbox)
        
        scroll.setWidget(scroll_content)
        exclusion_layout.addWidget(scroll)
        
        # Add tabs to tab widget
        tabs.addTab(main_tab, "Meal Planner")
        tabs.addTab(exclusion_tab, "Exclude Items")
        
        # Add tab widget to main layout
        layout.addWidget(tabs)
        
        # Initialize items list
        self.items = MEAL_ITEMS
        
        # Days of the week in Arabic
        self.days = [
            "السبت", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة",
            "السبت", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة"
        ]
        
        # Initialize the table with dropdowns
        self.initialize_table()

    def get_excluded_items(self):
        return [name for name, checkbox in self.exclusion_checkboxes.items() if checkbox.isChecked()]

    def initialize_table(self):
        self.table.setRowCount(len(self.days))
        
        # Get excluded items
        excluded_items = self.get_excluded_items()
        
        # Create dropdowns for each meal cell
        for row in range(len(self.days)):
            # Day column
            day_item = QTableWidgetItem(self.days[row])
            self.table.setItem(row, 0, day_item)
            
            # Breakfast column (Group 1 + Group 2)
            breakfast_combo = QComboBox()
            breakfast_group1_items = [item["name"] for item in self.items 
                                    if item["eat_time"] == "Breakfast" 
                                    and item["group"] == 1
                                    and item["name"] not in excluded_items]
            breakfast_group2_items = [item["name"] for item in self.items 
                                    if item["eat_time"] == "Breakfast" 
                                    and item["group"] == 2
                                    and item["name"] not in excluded_items]
            breakfast_combinations = [f"{g1} + {g2}" for g1 in breakfast_group1_items 
                                    for g2 in breakfast_group2_items]
            breakfast_combo.addItems(breakfast_combinations)
            # Set random value
            if breakfast_combinations:
                breakfast_combo.setCurrentIndex(random.randint(0, len(breakfast_combinations) - 1))
            self.table.setCellWidget(row, 1, breakfast_combo)
            
            # Lunch column (Group 1 + Group 2)
            lunch_combo = QComboBox()
            lunch_group1_items = [item["name"] for item in self.items 
                                if item["eat_time"] == "Lunch" 
                                and item["group"] == 1
                                and item["name"] not in excluded_items]
            lunch_group2_items = [item["name"] for item in self.items 
                                if item["eat_time"] == "Lunch" 
                                and item["group"] == 2
                                and item["name"] not in excluded_items]
            lunch_combinations = [f"{g1} + {g2}" for g1 in lunch_group1_items 
                                for g2 in lunch_group2_items]
            lunch_combo.addItems(lunch_combinations)
            # Set random value
            if lunch_combinations:
                lunch_combo.setCurrentIndex(random.randint(0, len(lunch_combinations) - 1))
            self.table.setCellWidget(row, 2, lunch_combo)
            
            # Dinner column (Group 1 only)
            dinner_combo = QComboBox()
            dinner_items = [item["name"] for item in self.items 
                          if item["eat_time"] == "Dinner" 
                          and item["group"] == 1
                          and item["name"] not in excluded_items]
            dinner_combo.addItems(dinner_items)
            # Set random value
            if dinner_items:
                dinner_combo.setCurrentIndex(random.randint(0, len(dinner_items) - 1))
            self.table.setCellWidget(row, 3, dinner_combo)
        
        # Enable save button after table is initialized
        self.save_button.setEnabled(True)

    def generate_meal_plan(self):
        # Get category counts from spin boxes
        category_a_count = self.category_a_spin.value()
        category_b_count = self.category_b_spin.value()
        category_c_count = self.category_c_spin.value()
        
        # Verify the sum is 7
        total_count = category_a_count + category_b_count + category_c_count
        if total_count != 7:
            QMessageBox.warning(self, "Error", "Category counts must sum to 7")
            return
        
        # Scale counts to sum to 14 (for 14 days)
        days = 14
        category_a_quota = round((category_a_count / total_count) * days)
        category_b_quota = round((category_b_count / total_count) * days)
        category_c_quota = round((category_c_count / total_count) * days)
        
        # Adjust quotas to ensure they sum to 14
        total_quota = category_a_quota + category_b_quota + category_c_quota
        if total_quota != days:
            diff = days - total_quota
            if category_a_quota >= category_b_quota and category_a_quota >= category_c_quota:
                category_a_quota += diff
            elif category_b_quota >= category_a_quota and category_b_quota >= category_c_quota:
                category_b_quota += diff
            else:
                category_c_quota += diff
        
        # Initialize the table with dropdowns
        self.initialize_table()

    def save_to_excel(self):
        try:
            file_name, _ = QFileDialog.getSaveFileName(
                self,
                "Save Excel File",
                "",
                "Excel Files (*.xlsx)"
            )
            
            if not file_name:
                return
                
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "خطة الوجبات"
            
            # Write headers
            headers = ["اليوم", "الإفطار", "الغداء", "العشاء"]
            for col, header in enumerate(headers, 1):
                sheet.cell(row=1, column=col).value = header
            
            # Get excluded items
            excluded_items = self.get_excluded_items()
            
            # Get all possible combinations for dropdowns
            # Breakfast combinations
            breakfast_group1_items = [item["name"] for item in self.items 
                                    if item["eat_time"] == "Breakfast" 
                                    and item["group"] == 1
                                    and item["name"] not in excluded_items]
            breakfast_group2_items = [item["name"] for item in self.items 
                                    if item["eat_time"] == "Breakfast" 
                                    and item["group"] == 2
                                    and item["name"] not in excluded_items]
            breakfast_combinations = [f"{g1} + {g2}" for g1 in breakfast_group1_items 
                                    for g2 in breakfast_group2_items]
            
            # Lunch combinations
            lunch_group1_items = [item["name"] for item in self.items 
                                if item["eat_time"] == "Lunch" 
                                and item["group"] == 1
                                and item["name"] not in excluded_items]
            lunch_group2_items = [item["name"] for item in self.items 
                                if item["eat_time"] == "Lunch" 
                                and item["group"] == 2
                                and item["name"] not in excluded_items]
            lunch_combinations = [f"{g1} + {g2}" for g1 in lunch_group1_items 
                                for g2 in lunch_group2_items]
            
            # Dinner items
            dinner_items = [item["name"] for item in self.items 
                          if item["eat_time"] == "Dinner" 
                          and item["group"] == 1
                          and item["name"] not in excluded_items]
            
            # Create named ranges for dropdowns
            # Breakfast
            for i, combo in enumerate(breakfast_combinations, 1):
                sheet.cell(row=i, column=6).value = combo
            breakfast_range = f"$F$1:$F${len(breakfast_combinations)}"
            
            # Lunch
            for i, combo in enumerate(lunch_combinations, 1):
                sheet.cell(row=i, column=7).value = combo
            lunch_range = f"$G$1:$G${len(lunch_combinations)}"
            
            # Dinner
            for i, item in enumerate(dinner_items, 1):
                sheet.cell(row=i, column=8).value = item
            dinner_range = f"$H$1:$H${len(dinner_items)}"
            
            # Write meal plan and create dropdown lists
            for row in range(self.table.rowCount()):
                # Day
                sheet.cell(row=row+2, column=1).value = self.table.item(row, 0).text()
                
                # Get selected values
                breakfast_combo = self.table.cellWidget(row, 1)
                lunch_combo = self.table.cellWidget(row, 2)
                dinner_combo = self.table.cellWidget(row, 3)
                
                # Write selected values
                sheet.cell(row=row+2, column=2).value = breakfast_combo.currentText()
                sheet.cell(row=row+2, column=3).value = lunch_combo.currentText()
                sheet.cell(row=row+2, column=4).value = dinner_combo.currentText()
                
                # Add data validation (dropdown lists)
                breakfast_dv = DataValidation(type="list", formula1=f"={breakfast_range}", allow_blank=True)
                lunch_dv = DataValidation(type="list", formula1=f"={lunch_range}", allow_blank=True)
                dinner_dv = DataValidation(type="list", formula1=f"={dinner_range}", allow_blank=True)
                
                # Apply data validation to cells
                sheet.add_data_validation(breakfast_dv)
                sheet.add_data_validation(lunch_dv)
                sheet.add_data_validation(dinner_dv)
                
                breakfast_dv.add(sheet.cell(row=row+2, column=2))
                lunch_dv.add(sheet.cell(row=row+2, column=3))
                dinner_dv.add(sheet.cell(row=row+2, column=4))
            
            # Hide the helper columns
            sheet.column_dimensions['F'].hidden = True
            sheet.column_dimensions['G'].hidden = True
            sheet.column_dimensions['H'].hidden = True
            
            # Save the file
            workbook.save(file_name)
            QMessageBox.information(self, "Success", f"Meal plan saved to {file_name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save Excel file: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MealPlanner()
    window.show()
    sys.exit(app.exec()) 