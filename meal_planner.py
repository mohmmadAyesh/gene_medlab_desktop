import sys
import random
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QSpinBox, QPushButton, 
                           QTableWidget, QTableWidgetItem, QMessageBox,
                           QGroupBox, QFileDialog, QScrollArea, QCheckBox,
                           QTabWidget, QComboBox, QHeaderView)
from PySide6.QtCore import Qt
import os
import tempfile
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import qn, OxmlElement
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from docx2pdf import convert
from docx import Document
from docx.oxml.shared import OxmlElement, qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PySide6.QtGui import QFont
from docx.shared import Inches, Pt;
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from meal_items import MEAL_ITEMS, DIABETES_EXCLUDED_FOODS, KIDNEY_EXCLUDED_FOODS
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml import parse_xml
def create_element(name):
    return OxmlElement(name)
def create_attribute(element,name,value):
    element.set(qn(name),value)
def create_dropdown_element(options, selected=None):
    sdt = OxmlElement('w:sdt')
    sdtPr = OxmlElement('w:sdtPr')
    ddl = OxmlElement('w:dropDownList')
    for option in options:
        li = OxmlElement('w:listItem')
        li.set(qn('w:displayText'), option)
        li.set(qn('w:value'), option)
        ddl.append(li)
    sdtPr.append(ddl)
    sdt.append(sdtPr)

    sdtContent = OxmlElement('w:sdtContent')
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = selected or (options[0] if options else '')
    r.append(t)
    p.append(r)
    sdtContent.append(p)
    sdt.append(sdtContent)
    return sdt

class MealPlanner(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Meal Planner")
        self.setMinimumSize(1200, 800)
        self.health_conditions = []
        
        # Set application font for Arabic text
        app_font = QFont("Arial")
        app_font.setPointSize(11)
        QApplication.setFont(app_font)
        
        # Set application-wide style
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
            }
            QGroupBox {
                background-color: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 12px;
                margin-top: 1.5em;
                padding: 15px;
                font-weight: bold;
                color: #111827;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px;
                color: #111827;
                font-size: 14px;
                background-color: #ffffff;
            }
            QPushButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 8px;
                font-weight: bold;
                font-size: 13px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #2563eb;
            }
            QPushButton:disabled {
                background-color: #9ca3af;
                color: #f3f4f6;
            }
            QSpinBox {
                padding: 8px;
                border: 2px solid #e5e7eb;
                border-radius: 8px;
                background-color: #ffffff;
                color: #111827;
                font-size: 13px;
                min-width: 80px;
            }
            QSpinBox:hover {
                border-color: #3b82f6;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                border: none;
                background-color: #f3f4f6;
                border-radius: 4px;
                margin: 1px;
            }
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {
                background-color: #e5e7eb;
            }
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 12px;
                gridline-color: #f3f4f6;
                color: #111827;
                font-size: 13px;
            }
            QTableWidget::item {
                padding: 12px;
                border-bottom: 1px solid #f3f4f6;
            }
            QTableWidget::item:selected {
                background-color: #dbeafe;
                color: #1e40af;
            }
            QHeaderView::section {
                background-color: #f9fafb;
                padding: 12px;
                border: none;
                border-bottom: 2px solid #e5e7eb;
                color: #111827;
                font-weight: bold;
                font-size: 13px;
            }
            QTabWidget::pane {
                border: 1px solid #e5e7eb;
                border-radius: 12px;
                background-color: #ffffff;
                top: -1px;
            }
            QTabBar::tab {
                background-color: #f9fafb;
                border: 1px solid #e5e7eb;
                padding: 10px 20px;
                margin-right: 4px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                color: #6b7280;
                font-size: 13px;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                border-bottom: none;
                color: #3b82f6;
                font-weight: bold;
            }
            QTabBar::tab:hover:!selected {
                color: #3b82f6;
            }
            QScrollArea {
                border: none;
                background-color: #ffffff;
            }
            QCheckBox {
                spacing: 8px;
                color: #111827;
                font-size: 13px;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
                border-radius: 6px;
            }
            QCheckBox::indicator:unchecked {
                border: 2px solid #e5e7eb;
                background-color: #ffffff;
            }
            QCheckBox::indicator:checked {
                background-color: #3b82f6;
                border: 2px solid #3b82f6;
            }
            QCheckBox::indicator:hover {
                border-color: #3b82f6;
            }
            QLabel {
                color: #111827;
                font-size: 13px;
                font-weight: 500;
            }
            QComboBox {
                padding: 8px 12px;
                border: 2px solid #e5e7eb;
                border-radius: 8px;
                background-color: #ffffff;
                color: #111827;
                min-width: 250px;
                font-size: 13px;
                selection-background-color: #dbeafe;
                selection-color: #1e40af;
                text-align: right;
            }
            QComboBox:hover {
                border-color: #3b82f6;
            }
            QComboBox::drop-down {
                border: none;
                width: 24px;
            }
            QComboBox::down-arrow {
                width: 12px;
                height: 12px;
                margin-right: 8px;
                image: none;
                border: none;
                background: none;
            }
            QComboBox::down-arrow:after {
                content: "";
                display: block;
                width: 0;
                height: 0;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #6b7280;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                background-color: #ffffff;
                selection-background-color: #dbeafe;
                selection-color: #1e40af;
                padding: 4px;
            }
            QScrollArea, QWidget#scrollContent {
                background-color: #ffffff;
                border: none;
            }
            QScrollArea {
                border: 1px solid #e5e7eb;
                border-radius: 12px;
            }
        """)
        
        # Initialize the main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Create tab widget with modern styling
        tabs = QTabWidget()
        tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #f5f5f5;
                border: 1px solid #e0e0e0;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: white;
                border-bottom: none;
            }
        """)
        
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
        
        # Update button layout with spacing
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        self.generate_button = QPushButton("Generate Meal Plan")
        self.generate_button.setStyleSheet("""
            QPushButton {
                background-color: #10b981;
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: 8px;
                font-weight: bold;
                font-size: 14px;
                min-width: 150px;
            }
            QPushButton:hover {
                background-color: #059669;
            }
        """)
        self.generate_button.clicked.connect(self.generate_meal_plan)
        self.save_excel_button = QPushButton("Save to Excel")
        self.save_excel_button.clicked.connect(self.save_to_excel)
        self.save_excel_button.setEnabled(False)
        self.export_to_word_template_button = QPushButton("Export to Word Template")
        self.export_to_word_template_button.clicked.connect(self.save_to_template_word)
        self.export_to_word_template_button.setEnabled(False)
        self.save_word_button = QPushButton("Save to Word")
        self.save_word_button.clicked.connect(self.save_to_word)
        self.save_word_button.setEnabled(False)
        self.save_pdf_button = QPushButton("Save to PDF")
        self.save_pdf_button.clicked.connect(self.save_to_pdf)
        self.save_pdf_button.setEnabled(False)
        self.save_gdocs_button = QPushButton("Save to Google Docs")
        self.save_gdocs_button.clicked.connect(self.save_to_gdoc)
        self.save_gdocs_button.setEnabled(False)
        button_layout.addWidget(self.generate_button)
        button_layout.addWidget(self.save_excel_button)
        button_layout.addWidget(self.save_word_button)
        button_layout.addWidget(self.save_pdf_button)
        button_layout.addWidget(self.save_gdocs_button)
        button_layout.addWidget(self.export_to_word_template_button)
        
        # Table widget for displaying meal plan
        table_group = QGroupBox("Meal Plan")
        table_layout = QVBoxLayout()
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["اليوم", "الإفطار", "الغداء", "العشاء"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                gridline-color: #e0e0e0;
            }
            QTableWidget::item {
                padding: 8px;
            }
            QHeaderView::section {
                background-color: #f5f5f5;
                padding: 8px;
                border: none;
                border-bottom: 1px solid #e0e0e0;
            }
        """)
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
        scroll_content.setObjectName("scrollContent")
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(10)
        scroll_layout.setContentsMargins(15, 15, 15, 15)
        
        # Create checkboxes for each meal item
        self.exclusion_checkboxes = {}
        for item in MEAL_ITEMS:
            checkbox = QCheckBox(item["name"])
            checkbox.setLayoutDirection(Qt.RightToLeft)
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

        # Set table properties for better Arabic text display
        self.table.setLayoutDirection(Qt.RightToLeft)
        header = self.table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignRight | Qt.AlignVCenter)
        
        # Set minimum row height for better readability
        self.table.verticalHeader().setDefaultSectionSize(50)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(True)
        self.table.setGridStyle(Qt.SolidLine)
        self.table.setStyleSheet(self.table.styleSheet() + """
            QTableWidget {
                gridline-color: #f3f4f6;
            }
            QTableWidget::item {
                padding: 12px;
                border-bottom: 1px solid #f3f4f6;
            }
            QTableWidget::item:alternate {
                background-color: #fafafa;
            }
        """)
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
        
        # Update ComboBox style to ensure text visibility
        combo_style = """
        QComboBox {
            padding: 10px 12px;
            border: 2px solid #e5e7eb;
            border-radius: 8px;
            background-color: #ffffff;
            color: #111827;
            min-height: 40px;
            min-width: 250px;
            font-size: 14px;
        }
        QComboBox QAbstractItemView {
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            background-color: #ffffff;
            selection-background-color: #dbeafe;
            selection-color: #1e40af;
            padding: 4px;
            font-size: 13px;
        }
        QComboBox QAbstractItemView::item {
            min-height: 36px;
            padding: 6px 8px;
        }
        QComboBox::drop-down{
            width: 24px;
            border: none;
        }
        QComboBox::down-arrow{
        image: none;
        width: 0;
        height: 0;
        margin-right: 8px;
        border-left: 5px solid transparent;
        border-right: 5px solid transparent;
        border-top: 5px solid #6b7280;
        }
        """
       
        # Create dropdowns for each meal cell
        for row in range(len(self.days)):
            # Day column
            day_item = QTableWidgetItem(self.days[row])
            day_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(row, 0, day_item)
            
            # Breakfast column
            breakfast_combo = QComboBox()
            breakfast_combo.setStyleSheet(combo_style)

            breakfast_combo.setLayoutDirection(Qt.RightToLeft)
            breakfast_combo.view().setLayoutDirection(Qt.RightToLeft)
            breakfast_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
            breakfast_combo.setMinimumContentsLength(1)
            breakfast_combo.setMinimumHeight(42)
            breakfast_combo.setEditable(False)
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
            if breakfast_combinations:
                breakfast_combo.setCurrentIndex(random.randint(0, len(breakfast_combinations) - 1))
            self.table.setCellWidget(row, 1, breakfast_combo)
            
            # Lunch column
            lunch_combo = QComboBox()
            lunch_combo.setStyleSheet(combo_style)
            lunch_combo.setLayoutDirection(Qt.RightToLeft)
            lunch_combo.view().setLayoutDirection(Qt.RightToLeft)
            lunch_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
            lunch_combo.setMinimumContentsLength(1)
            lunch_combo.setMinimumHeight(42)
            lunch_combo.setEditable(False)
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
            if lunch_combinations:
                lunch_combo.setCurrentIndex(random.randint(0, len(lunch_combinations) - 1))
            self.table.setCellWidget(row, 2, lunch_combo)
            
            # Dinner column
            dinner_combo = QComboBox()
            dinner_combo.setStyleSheet(combo_style)
            dinner_combo.setLayoutDirection(Qt.RightToLeft)
            dinner_combo.view().setLayoutDirection(Qt.RightToLeft)
            dinner_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
            dinner_combo.setMinimumContentsLength(1)
            dinner_combo.setMinimumHeight(15)
            dinner_combo.setEditable(False)
            
            dinner_items = [item["name"] for item in self.items 
                          if item["eat_time"] == "Dinner" 
                          and item["group"] == 1
                          and item["name"] not in excluded_items]
            dinner_combo.addItems(dinner_items)
            if dinner_items:
                dinner_combo.setCurrentIndex(random.randint(0, len(dinner_items) - 1))
            self.table.setCellWidget(row, 3, dinner_combo)

        # Adjust table column widths
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 100)  # Width for the day column

        # Enable save buttons after table is initialized
        self.save_excel_button.setEnabled(True)
        self.save_word_button.setEnabled(True)
        self.save_pdf_button.setEnabled(True)
        self.save_gdocs_button.setEnabled(True)
        self.export_to_word_template_button.setEnabled(True)

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
            
            # Enable both save buttons after successful generation
            self.save_excel_button.setEnabled(True)
            self.save_word_button.setEnabled(True)
            self.save_pdf_button.setEnabled(True)
            
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
    def save_to_word(self):
        try:
            file_name, _ = QFileDialog.getSaveFileName(self, "Save Word File", "", "Word Files (*.docx)")
            if not file_name:
                return
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
            dinner_items = [item["name"] for item in self.items 
                            if item["eat_time"] == "Dinner" 
                            and item["group"] == 1 
                            and item["name"] not in excluded_items]
            doc = Document()
            conditions = []
            if 1 in self.health_conditions:
                conditions.append("Healthy")
            if 2 in self.health_conditions:
                conditions.append("Diabetes")
            if 3 in self.health_conditions:
                conditions.append("Kidney Disease")
            doc.add_paragraph(f"Health Conditions: {', '.join(conditions)}")
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            headers = ["اليوم", "الإفطار", "الغداء", "العشاء"]
            for i, header in enumerate(headers):
                hdr_cells[i].text = header
                hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for row_idx in range(self.table.rowCount()):
                row_cells = table.add_row().cells
                day = self.table.item(row_idx, 0).text()
                row_cells[0].text = day
                row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                def create_dropdown(cell, options, selected=None):
                    sdt = OxmlElement('w:sdt')
                    sdtPr = OxmlElement('w:sdtPr')
                    ddl = OxmlElement('w:dropDownList')
                    for idx, option in enumerate(options):
                        li  = OxmlElement('w:listItem')
                        li.set(qn('w:displayText'),option)
                        li.set(qn('w:value'),str(idx))
                        ddl.append(li)
                    sdtPr.append(ddl)
                    sdt.append(sdtPr)
                    sdtContent = OxmlElement('w:sdtContent')
                    p = OxmlElement('w:p')
                    r = OxmlElement('w:r')
                    t = OxmlElement('w:t')
                    current_value = selected if selected else options[0] if options else ''
                    t.text = current_value
                    r.append(t)
                    p.append(r)
                    sdtContent.append(p)
                    cell._element.append(sdt)
                breakfast_combo = self.table.cellWidget(row_idx, 1).currentText()
                create_dropdown(row_cells[1], breakfast_combinations, breakfast_combo)
                lunch_combo = self.table.cellWidget(row_idx, 2).currentText()
                create_dropdown(row_cells[2], lunch_combinations, lunch_combo)
                dinner_combo = self.table.cellWidget(row_idx, 3).currentText()
                create_dropdown(row_cells[3], dinner_items, dinner_combo)
            doc.save(file_name)
            QMessageBox.information(self, "Success", "Meal plan saved successfully!")               
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save document: {str(e)}")
    # def save_to_pdf(self):
    #     try:
    #         # Create a new Word document (we'll use this as an intermediate step)
    #         doc = Document()
            
    #         # Add title
    #         title = doc.add_heading('خطة الوجبات الأسبوعية', level=1)
    #         title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
    #         # Create table
    #         table = doc.add_table(rows=1, cols=4)
    #         table.style = 'Table Grid'
            
    #         # Add headers
    #         headers = table.rows[0].cells
    #         headers[0].text = 'اليوم'
    #         headers[1].text = 'الإفطار'
    #         headers[2].text = 'الغداء'
    #         headers[3].text = 'العشاء'
            
    #         # Set header style
    #         for cell in headers:
    #             for paragraph in cell.paragraphs:
    #                 paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    #                 for run in paragraph.runs:
    #                     run.font.bold = True
    #                     run.font.size = Pt(12)
            
    #         # Add data rows from the table widget
    #         for row in range(self.table.rowCount()):
    #             cells = table.add_row().cells
    #             # Day
    #             cells[0].text = self.table.item(row, 0).text()
    #             # Meals - only get the current selection
    #             cells[1].text = self.table.cellWidget(row, 1).currentText()
    #             cells[2].text = self.table.cellWidget(row, 2).currentText()
    #             cells[3].text = self.table.cellWidget(row, 3).currentText()
                
    #             # Center align all cells
    #             for cell in cells:
    #                 for paragraph in cell.paragraphs:
    #                     paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
    #         # Set column widths
    #         for column in table.columns:
    #             for cell in column.cells:
    #                 cell.width = Inches(2.5)
            
    #         # Get save file path
    #         file_path, _ = QFileDialog.getSaveFileName(
    #             self,
    #             "Save PDF Document",
    #             "",
    #             "PDF Files (*.pdf)"
    #         )
            
    #         if not file_path:
    #             return
                
    #         # First save as docx
    #         with tempfile.TemporaryDirectory() as temp_dir:
    #             temp_docx  = os.path.join(temp_dir, "temp_meal_plan.docx")
    #             doc.save(temp_docx)
    #             convert(temp_docx, file_path)
    #         QMessageBox.information(self, "Success", "Meal plan saved to PDF successfully!")
    #     except Exception as e:
    #         QMessageBox.critical(self, "Error", f"Failed to save PDF: {str(e)}")
    def save_to_pdf(self):
        try:
        # Get save file path
            file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save PDF Document",
            "",
            "PDF Files (*.pdf)"
        )
        
            if not file_path:
                return

            from weasyprint import HTML, CSS
            from weasyprint.text.fonts import FontConfiguration
        
        # Create HTML content
            html_content = f"""
        <html dir="rtl">
        <head>
            <meta charset="UTF-8">
            <style>
                @page {{
                    size: A4;
                    margin: 1cm;
                }}
                body {{
                    font-family: Arial, sans-serif;
                    direction: rtl;
                }}
                h1 {{
                    text-align: center;
                    color: black;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 20px;
                }}
                th, td {{
                    border: 1px solid black;
                    padding: 8px;
                    text-align: center;
                }}
                th {{
                    background-color: #f2f2f2;
                    font-weight: bold;
                }}
            </style>
        </head>
        <body>
            <h1>خطة الوجبات الأسبوعية</h1>
            <table>
                <tr>
                    <th>اليوم</th>
                    <th>الإفطار</th>
                    <th>الغداء</th>
                    <th>العشاء</th>
                </tr>
        """
        
        # Add table rows
            for row in range(self.table.rowCount()):
                day = self.table.item(row, 0).text()
                breakfast = self.table.cellWidget(row, 1).currentText()
                lunch = self.table.cellWidget(row, 2).currentText()
                dinner = self.table.cellWidget(row, 3).currentText()
            
                html_content += f"""
                <tr>
                    <td>{day}</td>
                    <td>{breakfast}</td>
                    <td>{lunch}</td>
                    <td>{dinner}</td>
                </tr>
            """
        
            html_content += """
            </table>
        </body>
        </html>
        """
        
        # Configure fonts
            font_config = FontConfiguration()
        
        # Create PDF
            HTML(string=html_content).write_pdf(
            file_path,
            font_config=font_config,
            presentational_hints=True
            )
        
            QMessageBox.information(self, "Success", "Meal plan saved to PDF successfully!")
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save PDF: {str(e)}")
    def get_breakfast_combinations(self):
        excluded_items  = self.get_excluded_items();
        group1 = [item["name"] for item in self.items
                  if item["eat_time"] == "Breakfast" and item["group"] == 1 and item["name"] not in excluded_items]
        group2 = [item["name"] for item in self.items
                  if item["eat_time"] == "Breakfast" and item["group"] == 2 and item["name"] not in excluded_items]
        return [f"{g1} + {g2}" for g1 in group1 for g2 in group2]
    def get_lunch_combinations(self):
        excluded_items = self.get_excluded_items()
        group1 = [item["name"] for item in self.items
                  if item["eat_time"] == "Lunch" and item["group"] == 1 and item["name"] not in excluded_items]
        group2 = [item["name"] for item in self.items
                  if item["eat_time"] == "Lunch" and item["group"] == 2 and item["name"] not in excluded_items]
        return [f"{g1} + {g2}" for g1 in group1 for g2 in group2]
    def get_dinner_items(self):
        excluded_items = self.get_excluded_items()
        return [item["name"] for item in self.items
                if item["eat_time"] == "Dinner" and item["group"] == 1 and item["name"] not in excluded_items]
        
    def get_gdocs_service(self):
        SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive.file']
        flow = InstalledAppFlow.from_client_secrets_file("google_doc_credential.json", SCOPES)
        creds = flow.run_local_server(port=0)
        return build('docs', 'v1', credentials=creds)
    def get_health_conditions_text(self):
        conditions = []
        if hasattr(self,"health_conditions"):
            if 1 in self.health_conditions:
                conditions.append("Healthy")
            if 2 in self.health_conditions:
                conditions.append("Diabetes")
            if 3 in self.health_conditions:
                conditions.append("Kidney Disease")
        return conditions;
    def save_to_gdoc(self):
        try:
            service = self.get_gdocs_service()
            title = "Generated Meal Plan"
            doc = service.documents().create(body={"title": title}).execute()
            doc_id = doc['documentId']

            # Basic content
            requests = []
            requests.append({
                'insertText': {
                    'location': {'index': 1},
                    'text': f"Health Conditions: {', '.join(self.get_health_conditions_text())}\n\n"
                }
            })

            for row_idx in range(self.table.rowCount()):
                day = self.table.item(row_idx, 0).text()
                breakfast = self.table.cellWidget(row_idx, 1).currentText()
                lunch = self.table.cellWidget(row_idx, 2).currentText()
                dinner = self.table.cellWidget(row_idx, 3).currentText()

                for label, options, selected in [("الإفطار", self.get_breakfast_combinations(), breakfast),
                                             ("الغداء", self.get_lunch_combinations(), lunch),
                                             ("العشاء", self.get_dinner_items(), dinner)]:
                    requests.append({
                        "insertText": {
                            "location": {"index": 1},
                            "text": f"{day} - {label}: {selected}\n"
                        }
                    })

            service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()
            QMessageBox.information(self, "Success", f"Google Doc created: https://docs.google.com/document/d/{doc_id}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create Google Doc: {str(e)}")

    def save_to_template_word(self):
        try:
            # Step 1: Load template
            template_path = "num133333 (5).docx"
            if not os.path.exists(template_path):
                QMessageBox.critical(self, "Error", "Template file not found!")
                return
                
            doc = Document(template_path)
            
            # Step 2: Get meal options and current selections
            breakfast_options = self.get_breakfast_combinations()
            lunch_options = self.get_lunch_combinations()
            dinner_options = self.get_dinner_items()
            
            # Step 3: Create replacements dictionary with current selections
            replacements = {}
            for row in range(self.table.rowCount()):
                day_num = row + 1
                breakfast_key = f"Bf{day_num}"
                lunch_key = f"Lunch{day_num}"
                dinner_key = f"Dinner{day_num}"
                
                breakfast_selected = self.table.cellWidget(row, 1).currentText()
                lunch_selected = self.table.cellWidget(row, 2).currentText()
                dinner_selected = self.table.cellWidget(row, 3).currentText()
                
                replacements[breakfast_key] = (breakfast_options, breakfast_selected)
                replacements[lunch_key] = (lunch_options, lunch_selected)
                replacements[dinner_key] = (dinner_options, dinner_selected)

            # Step 4: Process shapes in the document
            for shape in doc.part.package.parts:
                if hasattr(shape, '_element') and hasattr(shape._element, 'txbx'):
                    try:
                        text_frame = shape._element.txbx
                        if text_frame is not None:
                            text_content = text_frame.text
                            
                            # Check for placeholders
                            for key, (options, selected) in replacements.items():
                                if key in text_content:
                                    # Create new paragraph with dropdown
                                    new_p = OxmlElement('w:p')
                                    dropdown = create_dropdown_element(options, selected)
                                    new_p.append(dropdown)
                                    
                                    # Replace text frame content
                                    text_frame.clear_content()
                                    text_frame._element.append(new_p)
                    except Exception as shape_error:
                        print(f"Error processing shape: {shape_error}")
                        continue

            # Alternative approach using word._element
            body = doc._element.body
            for shape in body.findall('.//w:txbxContent', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                for paragraph in shape.findall('.//w:p', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    text_content = ""
                    for run in paragraph.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        if run.text:
                            text_content += run.text
                    
                    # Check for placeholders
                    for key, (options, selected) in replacements.items():
                        if key in text_content:
                            # Clear paragraph content
                            for child in list(paragraph):
                                paragraph.remove(child)
                            
                            # Add dropdown
                            dropdown = create_dropdown_element(options, selected)
                            paragraph.append(dropdown)

            # Step 5: Save the document
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save File",
                "",
                "Word Files (*.docx)"
            )
            
            if save_path:
                doc.save(save_path)
                QMessageBox.information(self, "Success", "Meal plan saved with template successfully!")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save with template: {str(e)}")
            import traceback
            print(traceback.format_exc())
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MealPlanner()
    window.show()
    sys.exit(app.exec()) 