#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Aug  6 17:37:29 2021
Revised: 8/10/2021

@author: Don Spickler

This script sets up the options pane for the program.

"""

from PySide6.QtWidgets import (QSpinBox, QHBoxLayout,
                               QVBoxLayout, QWidget, QLabel, QGroupBox, QComboBox, QFormLayout, QCheckBox,
                               QRadioButton, QButtonGroup, QStackedWidget, QSizePolicy)


class OptionsEditorPane(QWidget):

    def __init__(self):
        super().__init__()
        self.initializeUI()

    def initializeUI(self):
        """
        Set up options widget.
        """

        # Make the Grid Type selections
        types = ["longtable", "tabular", "tabbing", "array", "matrix", "Special Matrix"]

        self.types_selector = QComboBox()
        self.types_selector.addItems(types)
        self.types_selector.currentIndexChanged.connect(self.typeChanged)

        # Create the separate option panels for each type.
        self.createTableOptions()
        self.createTabbingOptions()
        self.createArrayOptions()
        self.createMatrixOptions()
        self.createSpecialMatrixOptions()

        # Make the Includes selections
        self.include_math_mode = QCheckBox("Math Mode")
        self.include_array_stretch = QCheckBox("Array Stretch")

        include_layout = QVBoxLayout()
        include_layout.addWidget(self.include_math_mode)
        include_layout.addWidget(self.include_array_stretch)

        include_group = QGroupBox("Includes")
        include_group.setLayout(include_layout)

        grid_type_layout = QVBoxLayout()
        grid_type_layout.addWidget(self.types_selector)

        # Make Grid type selection.
        grid_type_group = QGroupBox("Grid Type")
        grid_type_group.setLayout(grid_type_layout)

        app_form_layout = QFormLayout()
        app_form_layout.addRow(grid_type_group)

        self.OptionsLabel = QLabel("Longtable/Tabular Options")
        #self.OptionsLabel.setStyleSheet("font-weight: bold;")
        self.OptionsLabel.setStyleSheet("text-decoration: underline;")

        # Place specific type panels into a stacked widget.
        self.options_stack = QStackedWidget()
        self.options_stack.addWidget(self.table_options_widget)
        self.options_stack.addWidget(self.tabbing_options_widget)
        self.options_stack.addWidget(self.array_options_widget)
        self.options_stack.addWidget(self.matrix_options_widget)
        self.options_stack.addWidget(self.SpecialMatrix_options_widget)

        # Put the selection panes together into one widget.
        app_form_layout.addRow(self.OptionsLabel)
        app_form_layout.addRow(self.options_stack)
        app_form_layout.addRow(include_group)

        pane_layout = QVBoxLayout()
        pane_layout.addLayout(app_form_layout)
        pane_layout.addStretch(1)

        self.setLayout(pane_layout)

    def typeChanged(self):
        """
        Process a type change when the user selects a different grid type.
        """
        gridtype = self.types_selector.currentText()

        # types = ["longtable", "tabular", "tabbing", "array", "matrix", "Special Matrix"]
        index = 0
        self.include_array_stretch.setVisible(True)
        if (gridtype == "longtable") or (gridtype == "tabular"):
            index = 0
            self.OptionsLabel.setText("Longtable/Tabular Options")
        elif (gridtype == "tabbing"):
            index = 1
            self.include_array_stretch.setVisible(False)
            self.OptionsLabel.setText("Tabbing Options")
        elif (gridtype == "array"):
            index = 2
            self.OptionsLabel.setText("Array Options")
        elif (gridtype == "matrix"):
            index = 3
            self.OptionsLabel.setText("Matrix Options")
        elif (gridtype == "Special Matrix"):
            index = 4
            self.OptionsLabel.setText("Special Matrix Options")

        count = self.options_stack.count()
        for i in range(count):
            widget = self.options_stack.widget(i)
            widget.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)

        self.options_stack.setCurrentIndex(index)
        self.options_stack.widget(index).setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        self.options_stack.adjustSize()
        self.adjustSize()

    def createTabbingOptions(self):
        """
        Create the options GUI for the tabbing grid type.
        """
        self.tabbibg_column_width = QSpinBox()
        self.tabbibg_column_width.setRange(0, 1000)
        self.tabbibg_column_width.setMinimumWidth(75)
        self.tabbibg_column_width.setValue(20)
        self.tabbibg_column_width.setSuffix(" pt")

        hbox = QHBoxLayout()
        hbox.addWidget(self.tabbibg_column_width)

        column_align_group = QGroupBox("Column Width")
        column_align_group.setLayout(hbox)

        self.tabbing_options_widget = QWidget()
        tabbing_options_widget_layout = QFormLayout()
        tabbing_options_widget_layout.addRow(column_align_group)

        self.tabbing_options_widget.setLayout(tabbing_options_widget_layout)

    def createMatrixOptions(self):
        """
        Create the options GUI for the matrix grid type.
        """

        # Make the Decoration selections
        self.matrix_Dec_None = QRadioButton("None")
        self.matrix_Dec_Paren = QRadioButton("( )")
        self.matrix_Dec_Bracket = QRadioButton("[ ]")
        self.matrix_Dec_Det = QRadioButton("| |")

        dec_bg = QButtonGroup(self)
        dec_bg.addButton(self.matrix_Dec_None)
        dec_bg.addButton(self.matrix_Dec_Paren)
        dec_bg.addButton(self.matrix_Dec_Bracket)
        dec_bg.addButton(self.matrix_Dec_Det)

        self.matrix_Dec_None.setChecked(True)

        dec_layout = QHBoxLayout()
        dec_layout.addWidget(self.matrix_Dec_None)
        dec_layout.addWidget(self.matrix_Dec_Paren)
        dec_layout.addWidget(self.matrix_Dec_Bracket)
        dec_layout.addWidget(self.matrix_Dec_Det)
        dec_layout.addStretch()

        dec_group = QGroupBox("Decoration")
        dec_group.setLayout(dec_layout)

        self.matrix_options_widget = QWidget()
        array_options_widget_layout = QFormLayout()
        array_options_widget_layout.addRow(dec_group)

        self.matrix_options_widget.setLayout(array_options_widget_layout)

    def createSpecialMatrixOptions(self):
        """
        Create the options GUI for the special matrix grid type.
        """

        # Make matrix type selections.
        self.SpecialMatrix_p = QRadioButton("pmatrix")
        self.SpecialMatrix_b = QRadioButton("bmatrix")
        self.SpecialMatrix_v = QRadioButton("vmatrix")
        self.SpecialMatrix_V = QRadioButton("Vmatrix")

        SM_bg = QButtonGroup(self)
        SM_bg.addButton(self.SpecialMatrix_p)
        SM_bg.addButton(self.SpecialMatrix_b)
        SM_bg.addButton(self.SpecialMatrix_v)
        SM_bg.addButton(self.SpecialMatrix_V)

        self.SpecialMatrix_p.setChecked(True)

        SM_layout = QVBoxLayout()
        SM_layout.addWidget(self.SpecialMatrix_p)
        SM_layout.addWidget(self.SpecialMatrix_b)
        SM_layout.addWidget(self.SpecialMatrix_v)
        SM_layout.addWidget(self.SpecialMatrix_V)
        SM_layout.addStretch()

        SM_group = QGroupBox("Matrix Type")
        SM_group.setLayout(SM_layout)

        self.SpecialMatrix_options_widget = QWidget()
        SM_options_widget_layout = QFormLayout()
        SM_options_widget_layout.addRow(SM_group)

        self.SpecialMatrix_options_widget.setLayout(SM_options_widget_layout)

    def createArrayOptions(self):
        """
        Create the options GUI for the array grid type.
        """

        # Make the Column Alignment selections
        self.array_column_align_left = QRadioButton("Left")
        self.array_column_align_center = QRadioButton("Center")
        self.array_column_align_right = QRadioButton("Right")

        column_align_bg = QButtonGroup(self)
        column_align_bg.addButton(self.array_column_align_left)
        column_align_bg.addButton(self.array_column_align_center)
        column_align_bg.addButton(self.array_column_align_right)

        self.array_column_align_left.setChecked(True)

        column_align_layout = QHBoxLayout()
        column_align_layout.addWidget(self.array_column_align_left)
        column_align_layout.addWidget(self.array_column_align_center)
        column_align_layout.addWidget(self.array_column_align_right)
        column_align_layout.addStretch()

        column_align_group = QGroupBox("Column Alignment")
        column_align_group.setLayout(column_align_layout)

        # Make the Border and Column Divisions selections
        self.array_table_border = QCheckBox("Table Border")
        self.array_first_row_division = QCheckBox("Division After First Row")
        self.array_all_row_division = QCheckBox("Division on All Rows")
        self.array_first_column_division = QCheckBox("Division After First Column")
        self.array_all_column_division = QCheckBox("Division After All Columns")

        border_layout = QVBoxLayout()
        border_layout.addWidget(self.array_table_border)
        border_layout.addWidget(self.array_first_row_division)
        border_layout.addWidget(self.array_all_row_division)
        border_layout.addWidget(self.array_first_column_division)
        border_layout.addWidget(self.array_all_column_division)

        border_type_group = QGroupBox("Border and Column Divisions")
        border_type_group.setLayout(border_layout)

        # Make the Decoration selections
        self.array_Dec_None = QRadioButton("None")
        self.array_Dec_Paren = QRadioButton("( )")
        self.array_Dec_Bracket = QRadioButton("[ ]")
        self.array_Dec_Det = QRadioButton("| |")

        dec_bg = QButtonGroup(self)
        dec_bg.addButton(self.array_Dec_None)
        dec_bg.addButton(self.array_Dec_Paren)
        dec_bg.addButton(self.array_Dec_Bracket)
        dec_bg.addButton(self.array_Dec_Det)

        self.array_Dec_None.setChecked(True)

        dec_layout = QHBoxLayout()
        dec_layout.addWidget(self.array_Dec_None)
        dec_layout.addWidget(self.array_Dec_Paren)
        dec_layout.addWidget(self.array_Dec_Bracket)
        dec_layout.addWidget(self.array_Dec_Det)
        dec_layout.addStretch()

        dec_group = QGroupBox("Decoration")
        dec_group.setLayout(dec_layout)

        self.array_options_widget = QWidget()
        array_options_widget_layout = QFormLayout()
        array_options_widget_layout.addRow(column_align_group)
        array_options_widget_layout.addRow(border_type_group)
        array_options_widget_layout.addRow(dec_group)

        self.array_options_widget.setLayout(array_options_widget_layout)

    def createTableOptions(self):
        """
        Create the options GUI for the tabular and long table grid types.
        """

        # Make the Column Alignment selections
        self.column_align_left = QRadioButton("Left")
        self.column_align_center = QRadioButton("Center")
        self.column_align_right = QRadioButton("Right")

        column_align_bg = QButtonGroup(self)
        column_align_bg.addButton(self.column_align_left)
        column_align_bg.addButton(self.column_align_center)
        column_align_bg.addButton(self.column_align_right)

        self.column_align_left.setChecked(True)

        column_align_layout = QHBoxLayout()
        column_align_layout.addWidget(self.column_align_left)
        column_align_layout.addWidget(self.column_align_center)
        column_align_layout.addWidget(self.column_align_right)
        column_align_layout.addStretch()

        column_align_group = QGroupBox("Column Alignment")
        column_align_group.setLayout(column_align_layout)

        # Make the Border and Column Divisions selections
        self.table_border = QCheckBox("Table Border")
        self.first_row_division = QCheckBox("Division After First Row")
        self.all_row_division = QCheckBox("Division on All Rows")
        self.first_column_division = QCheckBox("Division After First Column")
        self.all_column_division = QCheckBox("Division After All Columns")

        border_layout = QVBoxLayout()
        border_layout.addWidget(self.table_border)
        border_layout.addWidget(self.first_row_division)
        border_layout.addWidget(self.all_row_division)
        border_layout.addWidget(self.first_column_division)
        border_layout.addWidget(self.all_column_division)

        border_type_group = QGroupBox("Border and Column Divisions")
        border_type_group.setLayout(border_layout)

        # Make the Column Header selections
        self.column_header_left = QRadioButton("Left")
        self.column_header_center = QRadioButton("Center")
        self.column_header_right = QRadioButton("Right")

        column_header_bg = QButtonGroup(self)
        column_header_bg.addButton(self.column_header_left)
        column_header_bg.addButton(self.column_header_center)
        column_header_bg.addButton(self.column_header_right)

        self.column_header_left.setChecked(True)

        column_header_align_layout = QHBoxLayout()
        column_header_align_layout.addWidget(self.column_header_left)
        column_header_align_layout.addWidget(self.column_header_center)
        column_header_align_layout.addWidget(self.column_header_right)
        column_header_align_layout.addStretch()

        self.column_header_bold = QCheckBox("Bold")
        self.column_header_italic = QCheckBox("Italic")
        self.column_header_underline = QCheckBox("Underline")

        column_header_style_layout = QHBoxLayout()
        column_header_style_layout.addWidget(self.column_header_bold)
        column_header_style_layout.addWidget(self.column_header_italic)
        column_header_style_layout.addWidget(self.column_header_underline)
        column_header_style_layout.addStretch()

        self.column_header_rows = QSpinBox()
        self.column_header_rows.setRange(1, 10000)
        self.column_header_rows.setMinimumWidth(75)
        self.column_header_rows.setValue(1)

        column_header_rows_layout = QHBoxLayout()
        column_header_rows_layout.addWidget(QLabel("Number of Rows"))
        column_header_rows_layout.addWidget(self.column_header_rows)
        column_header_rows_layout.addStretch()

        column_header_layout = QVBoxLayout()
        column_header_layout.addLayout(column_header_align_layout)
        column_header_layout.addLayout(column_header_style_layout)
        column_header_layout.addLayout(column_header_rows_layout)

        self.column_header_group = QGroupBox("Column Header")
        self.column_header_group.setLayout(column_header_layout)
        self.column_header_group.setCheckable(True)
        self.column_header_group.setChecked(False)

        # Make the Row Header selections
        self.row_header_left = QRadioButton("Left")
        self.row_header_center = QRadioButton("Center")
        self.row_header_right = QRadioButton("Right")

        row_header_bg = QButtonGroup(self)
        row_header_bg.addButton(self.row_header_left)
        row_header_bg.addButton(self.row_header_center)
        row_header_bg.addButton(self.row_header_right)

        self.row_header_left.setChecked(True)

        row_header_align_layout = QHBoxLayout()
        row_header_align_layout.addWidget(self.row_header_left)
        row_header_align_layout.addWidget(self.row_header_center)
        row_header_align_layout.addWidget(self.row_header_right)
        row_header_align_layout.addStretch()

        self.row_header_bold = QCheckBox("Bold")
        self.row_header_italic = QCheckBox("Italic")
        self.row_header_underline = QCheckBox("Underline")

        row_header_style_layout = QHBoxLayout()
        row_header_style_layout.addWidget(self.row_header_bold)
        row_header_style_layout.addWidget(self.row_header_italic)
        row_header_style_layout.addWidget(self.row_header_underline)
        row_header_style_layout.addStretch()

        self.row_header_columns = QSpinBox()
        self.row_header_columns.setRange(1, 100)
        self.row_header_columns.setMinimumWidth(75)
        self.row_header_columns.setValue(1)

        row_header_columns_layout = QHBoxLayout()
        row_header_columns_layout.addWidget(QLabel("Number of Columns"))
        row_header_columns_layout.addWidget(self.row_header_columns)
        row_header_columns_layout.addStretch()

        row_header_layout = QVBoxLayout()
        row_header_layout.addLayout(row_header_align_layout)
        row_header_layout.addLayout(row_header_style_layout)
        row_header_layout.addLayout(row_header_columns_layout)

        self.row_header_group = QGroupBox("Row Header")
        self.row_header_group.setLayout(row_header_layout)
        self.row_header_group.setCheckable(True)
        self.row_header_group.setChecked(False)

        self.table_options_widget = QWidget()
        table_options_widget_layout = QFormLayout()
        table_options_widget_layout.addRow(column_align_group)
        table_options_widget_layout.addRow(border_type_group)
        table_options_widget_layout.addRow(self.column_header_group)
        table_options_widget_layout.addRow(self.row_header_group)

        self.table_options_widget.setLayout(table_options_widget_layout)

    def getOptionsInfo(self):
        """
        Create a dictionary of all options from the options panel.
        """

        info = {}
        info['Grid Type'] = self.types_selector.currentText()

        # Longtable/Tabular Options
        # Longtable/Tabular Alignment
        key = 'Table Column Align'
        if self.column_align_left.isChecked():
            info[key] = 'Left'
        elif self.column_align_center.isChecked():
            info[key] = 'Center'
        else:
            info[key] = 'Right'

        # Longtable/Tabular Border
        info['Table Border'] = self.table_border.isChecked()
        info['Table Division First Row'] = self.first_row_division.isChecked()
        info['Table Division All Rows'] = self.all_row_division.isChecked()
        info['Table Division First Column'] = self.first_column_division.isChecked()
        info['Table Division All Columns'] = self.all_column_division.isChecked()

        # Longtable/Tabular Column Headers
        info['Table Column Header'] = self.column_header_group.isChecked()
        key = 'Table Column Header Align'
        if self.column_header_left.isChecked():
            info[key] = 'Left'
        elif self.column_header_center.isChecked():
            info[key] = 'Center'
        else:
            info[key] = 'Right'

        info['Table Column Header Bold'] = self.column_header_bold.isChecked()
        info['Table Column Header Italic'] = self.column_header_italic.isChecked()
        info['Table Column Header Underline'] = self.column_header_underline.isChecked()
        info['Table Column Header Rows'] = self.column_header_rows.value()

        # Longtable/Tabular Row Headers
        info['Table Row Header'] = self.row_header_group.isChecked()
        key = 'Table Row Header Align'
        if self.row_header_left.isChecked():
            info[key] = 'Left'
        elif self.row_header_center.isChecked():
            info[key] = 'Center'
        else:
            info[key] = 'Right'

        info['Table Row Header Bold'] = self.row_header_bold.isChecked()
        info['Table Row Header Italic'] = self.row_header_italic.isChecked()
        info['Table Row Header Underline'] = self.row_header_underline.isChecked()
        info['Table Row Header Columns'] = self.row_header_columns.value()

        # Tabbing Options
        # Tabbing Column Width
        info['Tabbing Column Width'] = self.tabbibg_column_width.value()

        # Array Options
        # Array Alignment
        key = 'Array Column Align'
        if self.array_column_align_left.isChecked():
            info[key] = 'Left'
        elif self.array_column_align_center.isChecked():
            info[key] = 'Center'
        else:
            info[key] = 'Right'

        # Array Border
        info['Array Border'] = self.array_table_border.isChecked()
        info['Array Division First Row'] = self.array_first_row_division.isChecked()
        info['Array Division All Rows'] = self.array_all_row_division.isChecked()
        info['Array Division First Column'] = self.array_first_column_division.isChecked()
        info['Array Division All Columns'] = self.array_all_column_division.isChecked()

        # Array Decoration
        key = 'Array Decoration'
        if self.array_Dec_None.isChecked():
            info[key] = 'None'
        elif self.array_Dec_Paren.isChecked():
            info[key] = '()'
        elif self.array_Dec_Bracket.isChecked():
            info[key] = '[]'
        else:
            info[key] = '||'

        # Matrix Options
        key = 'Matrix Decoration'
        if self.matrix_Dec_None.isChecked():
            info[key] = 'None'
        elif self.matrix_Dec_Paren.isChecked():
            info[key] = '()'
        elif self.matrix_Dec_Bracket.isChecked():
            info[key] = '[]'
        else:
            info[key] = '||'

        # Special Matrix Options
        key = 'Special Matrix Decoration'
        if self.SpecialMatrix_p.isChecked():
            info[key] = 'pmatrix'
        elif self.SpecialMatrix_b.isChecked():
            info[key] = 'bmatrix'
        elif self.SpecialMatrix_v.isChecked():
            info[key] = 'vmatrix'
        else:
            info[key] = 'Vmatrix'

        # General Options
        info['Math Mode'] = self.include_math_mode.isChecked()
        info['Array Stretch'] = self.include_array_stretch.isChecked()

        # return dictionary
        return info
