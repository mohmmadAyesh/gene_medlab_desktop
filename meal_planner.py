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
from meal_items import MEAL_ITEMS, DIABETES_EXCLUDED_FOODS, KIDNEY_EXCLUDED_FOODS

class MealPlanner(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Meal Planner")
        self.setMinimumSize(1200, 800)
        self.health_conditions = []
        # Initialize the main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # Create tab widget
        tabs = QTabWidget()
        
        # Create main tab
        main_tab = QWidget()
        main_layout = QVBoxLayout(main_tab)
        #health tabls
        health_tab = QWidget()
        health_layout = QVBoxLayout(health_tab)
        health_group = QGroupBox("Health Conditions")
        health_group_layout = QVBoxLayout()
        self.healthy_checkbox = QCheckBox("Healthy (1)")
        self.diabetes_checkbox = QCheckBox("Diabetes (2)")
        self.kidney_checkbox = QCheckBox("Kidney Disease (3)")
        # Connect health condition checkboxes to update method
        self.healthy_checkbox.stateChanged.connect(self.update_health_conditions)
        self.diabetes_checkbox.stateChanged.connect(self.update_health_conditions)
        self.kidney_checkbox.stateChanged.connect(self.update_health_conditions)
        # add checkboxes to health group
        health_group_layout.addWidget(self.healthy_checkbox)
        health_group_layout.addWidget(self.diabetes_checkbox)
        health_group_layout.addWidget(self.kidney_checkbox)
        health_group.setLayout(health_group_layout)
        # Add health group to health tab
        health_layout.addWidget(health_group)
        health_layout.addStretch()
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
        tabs.addTab(health_tab, "Health Conditions")
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
    def update_health_conditions(self):
        """Update the health conditions array based on checkbox states"""
        self.health_conditions = []
        if self.healthy_checkbox.isChecked():
            self.health_conditions.append(1)
        if self.diabetes_checkbox.isChecked():
            self.health_conditions.append(2)
        if self.kidney_checkbox.isChecked():
            self.health_conditions.append(3)
        
        # Reinitialize table with updated conditions
        self.initialize_table()
    # def get_excluded_items(self):
    #     return [name for name, checkbox in self.exclusion_checkboxes.items() if checkbox.isChecked()]
    def get_excluded_items(self):
        """Get combined list of manually excluded items and condition-based exclusions"""
    # Get manually excluded items
        manual_exclusions = [name for name, checkbox in self.exclusion_checkboxes.items() 
                        if checkbox.isChecked()]
    
    # Add condition-based exclusions (now getting the names from the dictionaries)
        condition_exclusions = []
        if 2 in self.health_conditions:  # Diabetes
            condition_exclusions.extend([item["name"] for item in DIABETES_EXCLUDED_FOODS])
        if 3 in self.health_conditions:  # Kidney disease
            condition_exclusions.extend([item["name"] for item in KIDNEY_EXCLUDED_FOODS])
    
    # Return unique list of excluded item names
        all_exclusions = manual_exclusions + condition_exclusions
        return list(set(all_exclusions))
    def update_health_conditions(self):
        """Update the health conditions array based on checkbox states"""
        self.health_conditions = []
        if self.healthy_checkbox.isChecked():
            self.health_conditions.append(1)
        if self.diabetes_checkbox.isChecked():
            self.health_conditions.append(2)
        if self.kidney_checkbox.isChecked():
            self.health_conditions.append(3)
        
        # Reinitialize table with updated conditions
        self.initialize_table()
            

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
        try:
            # Get category counts
            count_a = self.category_a_spin.value()
            count_b = self.category_b_spin.value()
            count_c = self.category_c_spin.value()
            
            # Validate counts
            total = count_a + count_b + count_c
            if total != 7:
                QMessageBox.warning(self, "Error", "Category counts must sum to 7")
                return
            
            # Get excluded items
            excluded_items = self.get_excluded_items()
            
            # Get items for each category and meal time
            breakfast_group1_items = [item for item in self.items 
                                    if item["eat_time"] == "Breakfast" 
                                    and item["group"] == 1
                                    and item["name"] not in excluded_items]
            breakfast_group2_items = [item for item in self.items 
                                    if item["eat_time"] == "Breakfast" 
                                    and item["group"] == 2
                                    and item["name"] not in excluded_items]
            lunch_group1_items = [item for item in self.items 
                                if item["eat_time"] == "Lunch" 
                                and item["group"] == 1
                                and item["name"] not in excluded_items]
            lunch_group2_items = [item for item in self.items 
                                if item["eat_time"] == "Lunch" 
                                and item["group"] == 2
                                and item["name"] not in excluded_items]
            dinner_items = [item for item in self.items 
                          if item["eat_time"] == "Dinner" 
                          and item["group"] == 1
                          and item["name"] not in excluded_items]
            
            # Track previous day's meals
            prev_meals = {
                'breakfast_g1': [],
                'breakfast_g2': [],
                'lunch_g1': [],
                'lunch_g2': [],
                'dinner': []
            }
            
            # Generate meal plan
            for row in range(self.table.rowCount()):
                # Breakfast
                # Get available items excluding items from previous day
                available_breakfast_group1 = [item for item in breakfast_group1_items 
                                            if item not in prev_meals['breakfast_g1']]
                available_breakfast_group2 = [item for item in breakfast_group2_items 
                                            if item not in prev_meals['breakfast_g2']]
                
                # If no available items, use all items
                if not available_breakfast_group1:
                    available_breakfast_group1 = breakfast_group1_items
                if not available_breakfast_group2:
                    available_breakfast_group2 = breakfast_group2_items
                
                # Select random items with weighted probability
                # Give higher probability to items that haven't been used recently
                # Use a more aggressive weighting to prevent patterns
                weights_group1 = [1.0 if item not in prev_meals['breakfast_g1'] else 0.01 for item in available_breakfast_group1]
                weights_group2 = [1.0 if item not in prev_meals['breakfast_g2'] else 0.01 for item in available_breakfast_group2]
                
                # Add some randomness to the weights to prevent patterns
                weights_group1 = [w * random.uniform(0.8, 1.2) for w in weights_group1]
                weights_group2 = [w * random.uniform(0.8, 1.2) for w in weights_group2]
                
                # Normalize weights
                sum_weights1 = sum(weights_group1)
                sum_weights2 = sum(weights_group2)
                weights_group1 = [w/sum_weights1 for w in weights_group1]
                weights_group2 = [w/sum_weights2 for w in weights_group2]
                
                # Select items with weighted probability
                breakfast_group1 = random.choices(available_breakfast_group1, weights=weights_group1, k=1)[0]
                breakfast_group2 = random.choices(available_breakfast_group2, weights=weights_group2, k=1)[0]
                
                # Update previous meals tracking (keep only last day)
                prev_meals['breakfast_g1'] = [breakfast_group1]
                prev_meals['breakfast_g2'] = [breakfast_group2]
                
                # Lunch
                # Get available items excluding items from previous day
                available_lunch_group1 = [item for item in lunch_group1_items 
                                        if item not in prev_meals['lunch_g1']]
                available_lunch_group2 = [item for item in lunch_group2_items 
                                        if item not in prev_meals['lunch_g2']]
                
                # If no available items, use all items
                if not available_lunch_group1:
                    available_lunch_group1 = lunch_group1_items
                if not available_lunch_group2:
                    available_lunch_group2 = lunch_group2_items
                
                # Select random items
                lunch_group1 = random.choice(available_lunch_group1)
                lunch_group2 = random.choice(available_lunch_group2)
                
                # Update previous meals tracking (keep only last day)
                prev_meals['lunch_g1'] = [lunch_group1]
                prev_meals['lunch_g2'] = [lunch_group2]
                
                # Dinner
                # Get available items excluding items from previous day
                available_dinner = [item for item in dinner_items 
                                  if item not in prev_meals['dinner']]
                
                # If no available items, use all items
                if not available_dinner:
                    available_dinner = dinner_items
                
                # Select random item
                dinner = random.choice(available_dinner)
                
                # Update previous meals tracking (keep only last day)
                prev_meals['dinner'] = [dinner]
                
                # Create combinations
                breakfast_combo = f"{breakfast_group1['name']} + {breakfast_group2['name']}"
                lunch_combo = f"{lunch_group1['name']} + {lunch_group2['name']}"
                
                # Update table
                self.table.cellWidget(row, 1).setCurrentText(breakfast_combo)
                self.table.cellWidget(row, 2).setCurrentText(lunch_combo)
                self.table.cellWidget(row, 3).setCurrentText(dinner['name'])
            
            # Enable save button
            self.save_button.setEnabled(True)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate meal plan: {str(e)}")

    def save_to_excel(self):
        try:
            file_name,_ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
            if not file_name:
                return
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "خطة الوجبات"
            sheet.cell(row=1, column=1).value = "Health Conditions:"
            conditions = []
            if 1 in self.health_conditions:
                conditions.append("Healthy")
            if 2 in self.health_conditions:
                conditions.append("Diabetes")
            if 3 in self.health_conditions:
                conditions.append("Kidney Disease")
            sheet.cell(row=1, column=2).value = ", ".join(conditions)
            headers = ["اليوم", "الإفطار", "الغداء", "العشاء"]
            for col, header in enumerate(headers, 1):
                sheet.cell(row=3, column=col).value = header
            excluded_items = self.get_excluded_items()
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
            for i, combo in enumerate(breakfast_combinations, 1):
                sheet.cell(row=i, column=6).value = combo
            breakfast_range = f"$F$1:$F${len(breakfast_combinations)}"
            for i, combo in enumerate(lunch_combinations, 1):
                sheet.cell(row=i, column=7).value = combo
            lunch_range = f"$G$1:$G${len(lunch_combinations)}"
            for i, item in enumerate(dinner_items, 1):
                sheet.cell(row=i, column=8).value = item
            dinner_range = f"$H$1:$H${len(dinner_items)}"
            for row in range(self.table.rowCount()):
               sheet.cell(row=row+4, column=1).value = self.table.item(row, 0).text()
               breakfast_combo = self.table.cellWidget(row, 1)
               lunch_combo = self.table.cellWidget(row, 2)
               dinner_combo = self.table.cellWidget(row, 3)
               sheet.cell(row=row+4, column=2).value = breakfast_combo.currentText()
               sheet.cell(row=row+4, column=3).value = lunch_combo.currentText()
               sheet.cell(row=row+4, column=4).value = dinner_combo.currentText()
               breakfast_dv = DataValidation(type="list", formula1=f"={breakfast_range}", allow_blank=True)
               lunch_dv = DataValidation(type="list", formula1=f"={lunch_range}", allow_blank=True)
               dinner_dv = DataValidation(type="list", formula1=f"={dinner_range}", allow_blank=True)
               sheet.add_data_validation(breakfast_dv)
               sheet.add_data_validation(lunch_dv)
               sheet.add_data_validation(dinner_dv)
               breakfast_dv.add(sheet.cell(row=row+4, column=2))
               lunch_dv.add(sheet.cell(row=row+4, column=3))
               dinner_dv.add(sheet.cell(row=row+4, column=4))
            sheet.column_dimensions['F'].hidden = True
            sheet.column_dimensions['G'].hidden = True
            sheet.column_dimensions['H'].hidden = True
            workbook.save(file_name)
            QMessageBox.information(self, "Success", "Meal plan saved successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MealPlanner()
    window.show()
    sys.exit(app.exec()) 