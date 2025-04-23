import sys
import random
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QSpinBox, QPushButton, 
                           QTableWidget, QTableWidgetItem, QMessageBox,
                           QGroupBox, QFileDialog)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from meal_items import MEAL_ITEMS

class MealPlanner(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Meal Planner")
        self.setMinimumSize(1000, 800)
        
        # Initialize the main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
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
        
        # Add widgets to main layout
        layout.addWidget(input_group)
        layout.addLayout(button_layout)
        layout.addWidget(table_group)
        
        # Initialize items list
        self.items = MEAL_ITEMS
        
        # Days of the week in Arabic
        self.days = [
            "السبت", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة",
            "السبت", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة"
        ]

    def select_meal_items(self, meal_type, group, used_items, prev_day_items, category_counts, quotas):
        available_items = [item for item in self.items if item["eat_time"] == meal_type and item["group"] == group]
        
        if group == 1:
            available_items = [item for item in available_items if item["name"] not in prev_day_items]
            
            filtered_items = []
            for item in available_items:
                category = item["color"]
                if category == "A" and category_counts[meal_type]["A"] < quotas["A"]:
                    filtered_items.append(item)
                elif category == "B" and category_counts[meal_type]["B"] < quotas["B"]:
                    filtered_items.append(item)
                elif category == "C" and category_counts[meal_type]["C"] < quotas["C"]:
                    filtered_items.append(item)
            available_items = filtered_items
        
        if not available_items:
            available_items = [item for item in self.items if item["eat_time"] == meal_type and item["group"] == group]
            available_items = [item for item in available_items if item["name"] not in prev_day_items]
        
        if available_items:
            selected_item = random.choice(available_items)
            if group == 1:
                category_counts[meal_type][selected_item["color"]] += 1
            return selected_item["name"]
        return None

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
        
        # Initialize counters for category quotas per meal type
        category_counts = {
            "Breakfast": {"A": 0, "B": 0, "C": 0},
            "Lunch": {"A": 0, "B": 0, "C": 0},
            "Dinner": {"A": 0, "B": 0, "C": 0}
        }
        
        quotas = {
            "A": category_a_quota,
            "B": category_b_quota,
            "C": category_c_quota
        }
        
        # Clear and setup table
        self.table.setRowCount(len(self.days))
        
        used_breakfast_items = []
        used_lunch_items = []
        
        for i, day in enumerate(self.days):
            # Breakfast: Group 1 + Group 2
            prev_breakfast_items = used_breakfast_items[-1] if used_breakfast_items else []
            breakfast_group1 = self.select_meal_items("Breakfast", 1, used_breakfast_items, prev_breakfast_items, category_counts, quotas)
            breakfast_group2 = self.select_meal_items("Breakfast", 2, [], [], category_counts, quotas)
            breakfast_combo = f"{breakfast_group1} + {breakfast_group2}"
            used_breakfast_items.append([breakfast_group1])
            
            # Lunch: Group 1 + Group 2
            prev_lunch_items = used_lunch_items[-1] if used_lunch_items else []
            lunch_group1 = self.select_meal_items("Lunch", 1, used_lunch_items, prev_lunch_items, category_counts, quotas)
            lunch_group2 = self.select_meal_items("Lunch", 2, [], [], category_counts, quotas)
            lunch_combo = f"{lunch_group1} + {lunch_group2}"
            used_lunch_items.append([lunch_group1])
            
            # Dinner: Single Group 1 item
            dinner_item = self.select_meal_items("Dinner", 1, [], [], category_counts, quotas)
            
            # Add items to table
            self.table.setItem(i, 0, QTableWidgetItem(day))
            self.table.setItem(i, 1, QTableWidgetItem(breakfast_combo))
            self.table.setItem(i, 2, QTableWidgetItem(lunch_combo))
            self.table.setItem(i, 3, QTableWidgetItem(dinner_item))
        
        # Resize columns to fit content
        self.table.resizeColumnsToContents()
        self.save_button.setEnabled(True)

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
            
            # Write meal plan
            for row in range(self.table.rowCount()):
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    if item:
                        sheet.cell(row=row+2, column=col+1).value = item.text()
            
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