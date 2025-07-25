#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Aug  6 11:57:31 2021
Revised: 6/14/2025

@author: Don Spickler

This program allows the user to input a table into a standard spreadsheet
like grid and then export the LaTeX code to the clipboard given the options
the user selects.

"""

import pickle
import platform
import sys
import os

from PySide6.QtCore import Qt, QSize, QDir
from PySide6.QtGui import QIcon, QAction
from PySide6.QtWidgets import *

import webbrowser

# For the Mac OS
os.environ['QT_MAC_WANTS_LAYER'] = '1'


class LTCappcss:
    def __init__(self):
        super().__init__()
        self.css = """
            QMenu::separator { 
                background-color: #BBBBBB; 
                height: 1px; 
                margin: 2px 5px 2px 5px;
            }
            
            QMenu {padding: 4px 0px 4px 0px; }

        """

    def getCSS(self):
        return self.css

class LTC_Table(QTableWidget):
    def __init__(self, parent=None):
        super(LTC_Table, self).__init__(parent)
        self.setRowCount(3)
        self.setColumnCount(3)
        self.setSelectionMode(QAbstractItemView.ContiguousSelection)
        self.setCurrentCell(0, 0)
        self.cellChanged.connect(self.onCellChanged)
        self.tableHistory = []
        self.historyPos = 0
        self.addToHistory()
        ft = self.font()
        self.fontPointSize = 12
        ft.setPointSize(self.fontPointSize)
        self.setFont(ft)

    def keyPressEvent(self, event):
        """
        Overrides the key event processing.  Sends untracked keys to parent.
        """
        key = event.key()

        if key == Qt.Key_Return or key == Qt.Key_Enter:
            row = self.currentRow()
            col = self.currentColumn()
            row = row + 1
            if row >= self.rowCount():
                row = 0
                col = col + 1
            if col >= self.columnCount():
                col = 0
            self.setCurrentCell(row, col)
        elif key == Qt.Key_Delete:
            self.blockSignals(True)
            rng = self.selectedCellRanges()
            if len(rng) > 0:
                for i in range(rng[0][0], rng[0][1] + 1):
                    for j in range(rng[1][0], rng[1][1] + 1):
                        self.setItem(i, j, QTableWidgetItem(''))

            self.addToHistory()
            self.blockSignals(False)
        else:
            super(LTC_Table, self).keyPressEvent(event)

    def increaseFontSize(self):
        sz = self.fontPointSize
        if sz < 72:
            sz += 1
            self.adjustFontSize(sz)

    def decreaseFontSize(self):
        sz = self.fontPointSize
        if sz > 7:
            sz -= 1
            self.adjustFontSize(sz)

    def resetFontSize(self):
        self.adjustFontSize(12)

    def adjustFontSize(self, sz):
        ft = self.font()
        self.fontPointSize = sz
        ft.setPointSize(self.fontPointSize)
        self.setFont(ft)

    def addToHistory(self):
        """
        Adds the curent table to the undo/redo history.  Removes stored tables
        from the current table position to the end of the list.
        """
        currentTable = self.getTableContents()

        if len(self.tableHistory) > 0:
            while self.historyPos != len(self.tableHistory) - 1:
                self.tableHistory.pop(len(self.tableHistory) - 1)

        self.tableHistory.append(currentTable)
        self.historyPos = len(self.tableHistory) - 1
        self.setFocus()

    def onCellChanged(self):
        """
        if a cell is changed the new table is stored.
        """
        self.addToHistory()

    def selectedCellRanges(self):
        """
        Returns a lit of the upper left and lower right positions of the selected
        block of cells.
        """
        selitems = self.selectedIndexes()
        returnList = []
        if len(selitems) > 0:
            minrow = selitems[0].row()
            maxrow = selitems[0].row()
            mincol = selitems[0].column()
            maxcol = selitems[0].column()
            for i in selitems:
                if i.row() < minrow:
                    minrow = i.row()
                if i.row() > maxrow:
                    maxrow = i.row()
                if i.column() < mincol:
                    mincol = i.column()
                if i.column() > maxcol:
                    maxcol = i.column()

            returnList.append([minrow, maxrow])
            returnList.append([mincol, maxcol])
        return returnList

    def resizeTable(self, r, c):
        """
        Resizes the table to r rows and c columns.
        """
        self.blockSignals(True)
        self.setColumnCount(c)
        self.setRowCount(r)
        self.addToHistory()
        self.blockSignals(False)

    def closeEditing(self):
        """
        Stops the editing of a cell and takes the current edit as its contents.
        """
        # Trick to end editing of all cells, probably a better way to do this.
        self.setEnabled(False)
        self.setEnabled(True)

    def getTableContents(self):
        """
        Returns a list of row lists of table contents for the entire table.
        """
        self.closeEditing()
        tablelist = []
        for i in range(self.rowCount()):
            rowlist = []
            for j in range(self.columnCount()):
                item = self.item(i, j)
                if item == None:
                    rowlist.append('')
                else:
                    rowlist.append(item.text())
            tablelist.append(rowlist)
        return tablelist

    def getSelectedTableContents(self):
        """
        Returns a list of row lists of selected table contents.
        """
        self.closeEditing()
        tablelist = []
        rng = self.selectedCellRanges()
        if len(rng) > 0:
            for i in range(rng[0][0], rng[0][1] + 1):
                rowlist = []
                for j in range(rng[1][0], rng[1][1] + 1):
                    item = self.item(i, j)
                    if item == None:
                        rowlist.append('')
                    else:
                        rowlist.append(item.text())
                tablelist.append(rowlist)
        return tablelist

    def paste(self, items):
        """
        Pastes the list of row lists to the table, expanding the table size if
        necessary.
        """
        self.blockSignals(True)
        rows = len(items)
        cols = 0
        for i in range(rows):
            if len(items[i]) > cols:
                cols = len(items[i])

        rng = self.selectedCellRanges()
        if len(rng) > 0:
            startrow = rng[0][0]
            startcol = rng[1][0]

            if rows + startrow > self.rowCount():
                if rows + startrow > 10000:
                    rows = 10000 - startrow
                self.setRowCount(rows + startrow)

            if cols + startcol > self.columnCount():
                if cols + startcol > 1000:
                    cols = 1000 - startcol
                self.setColumnCount(cols + startcol)

            for i in range(startrow, rows + startrow):
                for j in range(startcol, cols + startcol):
                    self.setItem(i, j, QTableWidgetItem(items[i - startrow][j - startcol]))

        self.addToHistory()
        self.blockSignals(False)

    def getUpperLeftSelectedCell(self):
        """
        Returns the index of the upper left selected cell.
        """
        rng = self.selectedCellRanges()
        if len(rng) > 0:
            return [rng[0][0], rng[1][0]]
        else:
            return []

    def getLowerRightSelectedCell(self):
        """
        Returns the index of the lower right selected cell.
        """
        rng = self.selectedCellRanges()
        if len(rng) > 0:
            return [rng[0][1], rng[1][1]]
        else:
            return []

    def addRowAbove(self):
        """
        Adds a row above the selected cells.
        """
        start = self.getUpperLeftSelectedCell()
        if len(start) > 0:
            self.insertRow(start[0])
        self.addToHistory()

    def addRowBelow(self):
        """
        Adds a row below the selected cells.
        """
        start = self.getLowerRightSelectedCell()
        if len(start) > 0:
            self.insertRow(start[0] + 1)
        self.addToHistory()

    def addColumnBefore(self):
        """
        Adds a column before the selected cells.
        """
        start = self.getUpperLeftSelectedCell()
        if len(start) > 0:
            self.insertColumn(start[1])
        self.addToHistory()

    def addColumnAfter(self):
        """
        Adds a column after the selected cells.
        """
        start = self.getLowerRightSelectedCell()
        if len(start) > 0:
            self.insertColumn(start[1] + 1)
        self.addToHistory()

    def deleteRows(self):
        """
        Deletes the selected rows.
        """
        rng = self.selectedCellRanges()
        if len(rng) > 0:
            start = rng[0][0]
            end = rng[0][1]
            for i in range(end - start + 1):
                self.removeRow(start)
        if self.rowCount() == 0:
            self.insertRow(0)
        self.addToHistory()

    def deleteColumns(self):
        """
        Deletes the selected columns.
        """
        rng = self.selectedCellRanges()
        if len(rng) > 0:
            start = rng[1][0]
            end = rng[1][1]
            for i in range(end - start + 1):
                self.removeColumn(start)
        if self.columnCount() == 0:
            self.insertColumn(0)
        self.addToHistory()

    def deleteRowsColumns(self):
        """
        Deletes the selected rows and columns.
        """
        rng = self.selectedCellRanges()
        if len(rng) > 0:
            startr = rng[0][0]
            endr = rng[0][1]
            startc = rng[1][0]
            endc = rng[1][1]
            for i in range(endr - startr + 1):
                self.removeRow(startr)

            if self.rowCount() == 0:
                self.insertRow(0)

            for i in range(endc - startc + 1):
                self.removeColumn(startc)

            if self.columnCount() == 0:
                self.insertColumn(0)

        self.addToHistory()

    def transpose(self):
        """
        Transposes the grid.
        """
        self.blockSignals(True)

        items = self.getTableContents()
        rows = self.rowCount()
        cols = self.columnCount()

        tablelist = []
        for j in range(cols):
            collist = []
            for i in range(rows):
                item = items[i][j]
                if item == None:
                    collist.append('')
                else:
                    collist.append(item)
            tablelist.append(collist)

        self.setRowCount(1)
        self.setColumnCount(1)
        self.setCurrentCell(0, 0)
        self.paste(tablelist)

        self.setCurrentCell(0, 0)
        self.addToHistory()
        self.blockSignals(False)

    def trimcells(self):
        """
        Trims the entries in each cell.
        """
        self.blockSignals(True)
        for i in range(self.rowCount()):
            for j in range(self.columnCount()):
                item = self.item(i, j)

                if item == None:
                    item = ''
                else:
                    item = item.text()

                item = item.lstrip()
                item = item.rstrip()
                self.setItem(i, j, QTableWidgetItem(item))

        self.addToHistory()
        self.blockSignals(False)

    def fillcells(self, fill_text):
        """
        Fills the cells with the given text.
        """
        self.blockSignals(True)
        for i in range(self.rowCount()):
            for j in range(self.columnCount()):
                self.setItem(i, j, QTableWidgetItem(fill_text))

        self.addToHistory()
        self.blockSignals(False)

    def clearTable(self):
        """
        Clears the table contents.
        """
        self.clear()
        self.addToHistory()

    def newtable(self):
        """
        Clears the table contents and resets the size to 3 X 3.
        """
        self.setRowCount(3)
        self.setColumnCount(3)
        self.clear()
        self.setCurrentCell(0, 0)
        self.addToHistory()

    def loadItems(self, currentTable):
        """
        Loads the items (list of row lists) into the table.
        """
        self.blockSignals(True)
        rows = len(currentTable)
        cols = len(currentTable[0])
        self.setRowCount(rows)
        self.setColumnCount(cols)

        for i in range(self.rowCount()):
            for j in range(self.columnCount()):
                self.setItem(i, j, QTableWidgetItem(currentTable[i][j]))
        self.blockSignals(False)

    def undo(self):
        """
        Processes an undo.
        """
        self.historyPos = self.historyPos - 1
        if self.historyPos < 0:
            self.historyPos = 0

        currentTable = self.tableHistory[self.historyPos]
        self.loadItems(currentTable)

    def redo(self):
        """
        Processes a redo.
        """
        self.historyPos = self.historyPos + 1
        if self.historyPos >= len(self.tableHistory):
            self.historyPos = len(self.tableHistory) - 1

        currentTable = self.tableHistory[self.historyPos]
        self.loadItems(currentTable)

class LTCOptionsEditorPane(QWidget):

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


class LaTeXTableEditor(QMainWindow):

    def __init__(self, parent=None):
        """
        Initialize the program and set up the programList structure that
        allows the creation of child applications.
        """
        super().__init__()
        self.Parent = parent
        
        self.authors = "Don Spickler"
        self.version = "2.6.1"
        self.program_title = "Latex Table Creator"
        self.copyright = "2022 - 2025"
        self.licence = "\nThis software is distributed under the GNU General Public License version 3.\n\n" + \
        "This program is free software: you can redistribute it and/or modify it under the terms of the GNU " + \
        "General Public License as published by the Free Software Foundation, either version 3 of the License, " + \
        "or (at your option) any later version. This program is distributed in the hope that it will be useful, " + \
        "but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A " + \
        "PARTICULAR PURPOSE. See the GNU General Public License for more details http://www.gnu.org/licenses/."

        self.clipboard = QApplication.clipboard()
        self.programList = []
        self.setMinimumSize(800, 600)
        self.setWindowTitle('LaTeX Table Creator')
        icon = QIcon(self.resource_path("icons/ProgramIcon2.png"))
        self.setWindowIcon(icon)

        self.currentTheme = ''
        self.Platform = platform.system()
        styles = QStyleFactory.keys()
        if "Fusion" in styles:
            app.setStyle('Fusion')
            self.currentTheme = 'Fusion'
        else:
            self.currentTheme = styles[0]

        try:
            with open('LaTeXTableCreatorOptions.opt', 'rb') as f:
                filecontents = pickle.load(f)
                theme = filecontents[0]
                self.Parent.setStyle(theme)
                self.currentTheme = theme
        except Exception as e:
            pass

        self.createTablePane()
        self.createMenu()
        self.createToolBar()
        self.createDockWidget()
        self.show()

    def resource_path(self, relative_path):
        """
        Creates a system path that is rlative to the position of the running application.

        :param relative_path: The relative path of the file from the base position of the running appliction.
        :return: The full OS path.
        """
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)

    # def resource_path(self, relative_path):
    #     if hasattr(sys, '_MEIPASS'):
    #         return os.path.join(sys._MEIPASS, relative_path)
    #     return os.path.join(os.path.abspath("LaTeXTableCreator"), relative_path)

    def createTablePane(self):
        """
        Sets up the portion of the screen that contains the table and the
        size selection spinners.
        """
        self.table_widget = LTC_Table()

        row_label = QLabel("Rows")
        self.rows = QSpinBox()
        self.rows.setRange(1, 10000)
        self.rows.setMinimumWidth(75)
        self.rows.setValue(3)
        self.rows.valueChanged.connect(self.resizeTableRows)

        column_label = QLabel("Columns")
        self.columns = QSpinBox()
        self.columns.setRange(1, 1000)
        self.columns.setMinimumWidth(75)
        self.columns.setValue(3)
        self.columns.valueChanged.connect(self.resizeTableColumns)

        h_box_tablesize = QHBoxLayout()
        h_box_tablesize.addWidget(row_label)
        h_box_tablesize.addWidget(self.rows)
        h_box_tablesize.addWidget(column_label)
        h_box_tablesize.addWidget(self.columns)
        h_box_tablesize.addStretch()

        h_box_table = QHBoxLayout()
        h_box_table.addWidget(self.table_widget)

        v_box_tablepane = QVBoxLayout()
        v_box_tablepane.setContentsMargins(5, 5, 0, 0)
        v_box_tablepane.addLayout(h_box_tablesize)
        v_box_tablepane.addLayout(h_box_table)

        centerPane = QWidget()
        centerPane.setLayout(v_box_tablepane)
        self.setCentralWidget(centerPane)

    def resizeTableRows(self):
        """
        Resizes the table to the current spinner values.
        """
        self.resizeTable()
        self.rows.setFocus()

    def resizeTableColumns(self):
        """
        Resizes the table to the current spinner values.
        """
        self.resizeTable()
        self.columns.setFocus()

    def resizeTable(self):
        """
        Resizes the table to the current spinner values.
        """
        rows = self.rows.value()
        cols = self.columns.value()
        self.table_widget.resizeTable(rows, cols)

    def createMenu(self):
        """
        Set up the menu bar.
        """
        # Create file menu actions
        self.file_new_act = QAction(QIcon(self.resource_path('icons/FileNew.png')), "New", self)
        self.file_new_act.setShortcut('Ctrl+N')
        self.file_new_act.setStatusTip('Start a new table.')
        self.file_new_act.triggered.connect(self.newtable)

        self.file_open_act = QAction(QIcon(self.resource_path('icons/FileOpen.png')), "Open...", self)
        self.file_open_act.setShortcut('Ctrl+O')
        self.file_open_act.setStatusTip('Open a table file.')
        self.file_open_act.triggered.connect(self.openFile)

        self.file_saveas_act = QAction(QIcon(self.resource_path('icons/FileSave.png')), "Save As...", self)
        self.file_saveas_act.setShortcut('Ctrl+S')
        self.file_saveas_act.setStatusTip('Save a table file.')
        self.file_saveas_act.triggered.connect(self.saveFile)

        quit_act = QAction("Exit", self)
        quit_act.setStatusTip('Exit the program.')
        quit_act.triggered.connect(self.close)

        # Create table menu actions
        self.add_row_above_act = QAction("Add Row Above", self)
        self.add_row_above_act.setShortcut('Ctrl+Up')
        self.add_row_above_act.setStatusTip('Add a row above the current selection.')
        self.add_row_above_act.triggered.connect(self.addRowAbove)

        self.add_row_below_act = QAction("Add Row Below", self)
        self.add_row_below_act.setShortcut('Ctrl+Down')
        self.add_row_below_act.setStatusTip('Add a row below the current selection.')
        self.add_row_below_act.triggered.connect(self.addRowBelow)

        self.add_col_before_act = QAction("Add Column Before", self)
        self.add_col_before_act.setShortcut('Ctrl+Left')
        self.add_col_before_act.setStatusTip('Add a column before the current selection.')
        self.add_col_before_act.triggered.connect(self.addColumnBefore)

        self.add_col_after_act = QAction("Add Column After", self)
        self.add_col_after_act.setShortcut('Ctrl+Right')
        self.add_col_after_act.setStatusTip('Add a column after the current selection.')
        self.add_col_after_act.triggered.connect(self.addColumnAfter)

        self.delete_row_act = QAction("Delete Rows", self)
        self.delete_row_act.setStatusTip('Delete selected rows.')
        self.delete_row_act.triggered.connect(self.deleteRows)

        self.delete_col_act = QAction("Delete Columns", self)
        self.delete_col_act.setStatusTip('Delete selected columns.')
        self.delete_col_act.triggered.connect(self.deleteColumns)

        self.delete_row_col_act = QAction("Delete Rows and Columns", self)
        self.delete_row_col_act.setStatusTip('Delete selected rows and columns.')
        self.delete_row_col_act.triggered.connect(self.deleteRowsColumns)

        self.clear_table_act = QAction("Clear All", self)
        self.clear_table_act.setStatusTip('Delete the table contents.')
        self.clear_table_act.triggered.connect(self.clearTable)

        self.transpose_act = QAction("Transpose Table", self)
        self.transpose_act.setStatusTip('Transposes the table contents.')
        self.transpose_act.triggered.connect(self.transpose)

        self.trim_act = QAction("Trim Cells", self)
        self.trim_act.setStatusTip('Trim each of the cells.')
        self.trim_act.triggered.connect(self.trim)

        self.fill_cells_act = QAction("Fill Cells with Text...", self)
        self.fill_cells_act.setStatusTip('Fill each cell with the same value.')
        self.fill_cells_act.triggered.connect(self.filltext)

        self.adjust_widths_act = QAction(QIcon(self.resource_path('icons/AdjCol.png')), "Adjust Column Widths", self)
        self.adjust_widths_act.setStatusTip('Adjust the column widths to fit the contents.')
        self.adjust_widths_act.triggered.connect(self.adjustWidths)

        self.adjust_heights_act = QAction(QIcon(self.resource_path('icons/AdjRow.png')), "Adjust Row Heights", self)
        self.adjust_heights_act.setStatusTip('Adjust the row heights to fit the contents.')
        self.adjust_heights_act.triggered.connect(self.adjustHeights)

        self.adjust_width_height_act = QAction(QIcon(self.resource_path('icons/AdjRowCol.png')), "Adjust Row and Column Sizes", self)
        self.adjust_width_height_act.setShortcut('Ctrl+R')
        self.adjust_width_height_act.setStatusTip('Adjust the row and column sizes to fit the contents.')
        self.adjust_width_height_act.triggered.connect(self.adjustWidthHeight)

        # Create edit menu actions
        self.copy_selected_act = QAction(QIcon(self.resource_path('icons/Cascade.png')), "Copy Selected", self)
        self.copy_selected_act.setShortcut('Ctrl+C')
        self.copy_selected_act.setStatusTip('Copy table selection to the clipboard.')
        self.copy_selected_act.triggered.connect(self.copySelected)

        self.copy_all_act = QAction(QIcon(self.resource_path('icons/copy.png')), "Copy All", self)
        self.copy_all_act.setShortcut('Shift+Ctrl+C')
        self.copy_all_act.setStatusTip('Copy entire table to the clipboard.')
        self.copy_all_act.triggered.connect(self.copyAll)

        self.paste_act = QAction(QIcon(self.resource_path('icons/paste.png')), "Paste", self)
        self.paste_act.setShortcut('Ctrl+V')
        self.paste_act.setStatusTip('Paste clipboard to table.')
        self.paste_act.triggered.connect(self.paste)

        self.copy_latex_act = QAction(QIcon(self.resource_path('icons/latexicon2.png')), "Copy as LaTeX", self)
        self.copy_latex_act.setShortcut('Ctrl+L')
        self.copy_latex_act.setStatusTip('Copy table to LaTeX code with selected options.')
        self.copy_latex_act.triggered.connect(self.latexCopy)

        self.paste_from_latex_act = QAction("Paste from LaTeX", self)
        self.paste_from_latex_act.setShortcut('Shift+Ctrl+L')
        self.paste_from_latex_act.setStatusTip('Paste from LaTeX code to table.')
        self.paste_from_latex_act.triggered.connect(self.pasteLatex)

        self.copy_maxima_act = QAction(QIcon(self.resource_path('icons/wxmaximaicon002.png')),"Copy as Maxima", self)
        self.copy_maxima_act.setStatusTip('Copy matrix to Maxima code.')
        self.copy_maxima_act.triggered.connect(self.copyMaxima)

        self.copy_sage_act = QAction(QIcon(self.resource_path('icons/sagemath.png')),"Copy as SageMath", self)
        self.copy_sage_act.setStatusTip('Copy matrix to SageMath code.')
        self.copy_sage_act.triggered.connect(self.copySage)

        self.copy_geogebra_act = QAction(QIcon(self.resource_path('icons/Geogebra002.png')),"Copy as GeoGebra", self)
        self.copy_geogebra_act.setStatusTip('Copy matrix to Mathematica code.')
        self.copy_geogebra_act.triggered.connect(self.copyGeoGebra)

        self.copy_bracket_act = QAction("Copy [...] Delimited", self)
        self.copy_bracket_act.setStatusTip('Copy matrix in [...] delimited form.')
        self.copy_bracket_act.triggered.connect(self.copyBracket)

        self.copy_curleybracket_act = QAction("Copy {...} Delimited", self)
        self.copy_curleybracket_act.setStatusTip('Copy matrix in {...} delimited form.')
        self.copy_curleybracket_act.triggered.connect(self.copyGeoGebra)

        self.copy_angle_bracket_act = QAction("Copy <...> Delimited", self)
        self.copy_angle_bracket_act.setStatusTip('Copy matrix in <...> delimited form.')
        self.copy_angle_bracket_act.triggered.connect(self.copyAngleBracket)

        self.copy_html_act = QAction("Copy as HTML", self)
        self.copy_html_act.setStatusTip('Copy table to HTML.')
        self.copy_html_act.triggered.connect(self.copyHTML)

        self.select_all_act = QAction("Select All", self)
        self.select_all_act.setShortcut('Ctrl+A')
        self.select_all_act.setStatusTip('Select all cells.')
        self.select_all_act.triggered.connect(self.selectall)

        self.Undo_act = QAction(QIcon(self.resource_path('icons/Undo.png')), "Undo", self)
        self.Undo_act.setShortcut('Ctrl+Z')
        self.Undo_act.setStatusTip('Undo the last edit.')
        self.Undo_act.triggered.connect(self.undo)

        self.Redo_act = QAction(QIcon(self.resource_path('icons/Redo.png')), "Redo", self)
        self.Redo_act.setShortcut('Shift+Ctrl+Z')
        self.Redo_act.setStatusTip('Redo the last edit.')
        self.Redo_act.triggered.connect(self.redo)

        self.view_increase_font_size_act = QAction(QIcon(self.resource_path('icons/zoomin.png')), "Increase Font Size", self)
        self.view_increase_font_size_act.setStatusTip('Increase the font size in the workspace.')
        self.view_increase_font_size_act.triggered.connect(self.increaseWorksheetFontSize)

        self.view_decrease_font_size_act = QAction(QIcon(self.resource_path('icons/zoomout.png')), "Decrease Font Size", self)
        self.view_decrease_font_size_act.setStatusTip('Decrease the font size in the workspace.')
        self.view_decrease_font_size_act.triggered.connect(self.decreaseWorksheetFontSize)

        self.view_reset_font_size_act = QAction(QIcon(self.resource_path('icons/Tile.png')), "Reset Font Size", self)
        self.view_reset_font_size_act.setStatusTip('Reset the font size in the workspace to the defaults.')
        self.view_reset_font_size_act.triggered.connect(self.resetWorksheetFontSize)

        self.SelectTheme_act = QAction("Select Theme...", self)
        self.SelectTheme_act.triggered.connect(self.SelectTheme)
        self.SelectTheme_act.setStatusTip('Select from the current supported system themes.')

        # Create help menu actions
        self.help_about_act = QAction(QIcon(self.resource_path('icons/About.png')), "About...", self)
        self.help_about_act.setStatusTip('About the LaTeX Table Creator')
        self.help_about_act.triggered.connect(self.aboutDialog)

        self.help_help_act = QAction(QIcon(self.resource_path('icons/Help2.png')), "Help...", self)
        self.help_help_act.setStatusTip('Help with ' + self.program_title + " Version " + self.version + "...")
        self.help_help_act.triggered.connect(self.onHelp)

        # Create the menu bar
        menu_bar = self.menuBar()
        menu_bar.setNativeMenuBar(False)

        # Create file menu and add actions
        file_menu = menu_bar.addMenu('File')
        file_menu.addAction(self.file_new_act)
        file_menu.addAction(self.file_open_act)
        file_menu.addSeparator()
        file_menu.addAction(self.file_saveas_act)
        file_menu.addSeparator()
        file_menu.addAction(quit_act)

        # Create Edit menu and add actions
        edit_menu = menu_bar.addMenu('Edit')
        edit_menu.addAction(self.copy_selected_act)
        edit_menu.addAction(self.copy_all_act)
        edit_menu.addAction(self.select_all_act)
        edit_menu.addAction(self.paste_act)
        edit_menu.addSeparator()
        edit_menu.addAction(self.copy_latex_act)
        edit_menu.addAction(self.paste_from_latex_act)
        edit_menu.addSeparator()
        edit_menu.addAction(self.copy_geogebra_act)
        edit_menu.addAction(self.copy_maxima_act)
        edit_menu.addAction(self.copy_sage_act)
        edit_menu.addSeparator()
        edit_menu.addAction(self.copy_bracket_act)
        edit_menu.addAction(self.copy_curleybracket_act)
        edit_menu.addAction(self.copy_angle_bracket_act)
        edit_menu.addAction(self.copy_html_act)
        edit_menu.addSeparator()
        edit_menu.addAction(self.SelectTheme_act)
        edit_menu.addSeparator()
        edit_menu.addAction(self.Undo_act)
        edit_menu.addAction(self.Redo_act)

        # Create table menu and add actions
        table_menu = menu_bar.addMenu('Table')
        table_menu.addAction(self.add_row_above_act)
        table_menu.addAction(self.add_row_below_act)
        table_menu.addSeparator()
        table_menu.addAction(self.add_col_before_act)
        table_menu.addAction(self.add_col_after_act)
        table_menu.addSeparator()
        table_menu.addAction(self.delete_row_act)
        table_menu.addAction(self.delete_col_act)
        table_menu.addAction(self.delete_row_col_act)
        table_menu.addSeparator()
        table_menu.addAction(self.transpose_act)
        table_menu.addAction(self.trim_act)
        table_menu.addAction(self.fill_cells_act)
        table_menu.addSeparator()
        table_menu.addAction(self.clear_table_act)

        view_menu = menu_bar.addMenu('View')
        view_menu.addAction(self.adjust_widths_act)
        view_menu.addAction(self.adjust_heights_act)
        view_menu.addAction(self.adjust_width_height_act)
        view_menu.addSeparator()
        view_menu.addAction(self.view_increase_font_size_act)
        view_menu.addAction(self.view_decrease_font_size_act)
        view_menu.addAction(self.view_reset_font_size_act)

        # Create help menu and add actions
        help_menu = menu_bar.addMenu('Help')
        help_menu.addAction(self.help_help_act)
        help_menu.addAction(self.help_about_act)

    def createToolBar(self):
        """
        Create toolbar for GUI
        """
        # Set up toolbar
        tool_bar = QToolBar("Main Toolbar")
        tool_bar.setIconSize(QSize(18, 18))
        tool_bar.setMovable(False)
        self.addToolBar(tool_bar)

        # Add actions to toolbar
        tool_bar.addAction(self.file_new_act)
        tool_bar.addAction(self.file_open_act)
        tool_bar.addAction(self.file_saveas_act)
        tool_bar.addSeparator()
        tool_bar.addAction(self.copy_selected_act)
        tool_bar.addAction(self.copy_all_act)
        tool_bar.addAction(self.paste_act)
        tool_bar.addSeparator()
        tool_bar.addAction(self.copy_latex_act)
        tool_bar.addSeparator()
        tool_bar.addAction(self.adjust_widths_act)
        tool_bar.addAction(self.adjust_heights_act)
        tool_bar.addAction(self.adjust_width_height_act)
        tool_bar.addSeparator()
        tool_bar.addAction(self.view_increase_font_size_act)
        tool_bar.addAction(self.view_decrease_font_size_act)
        tool_bar.addAction(self.view_reset_font_size_act)
        tool_bar.addSeparator()
        tool_bar.addAction(self.Undo_act)
        tool_bar.addAction(self.Redo_act)
        tool_bar.addSeparator()
        tool_bar.addAction(self.help_help_act)
        tool_bar.addAction(self.help_about_act)

    def increaseWorksheetFontSize(self):
        self.table_widget.increaseFontSize()

    def decreaseWorksheetFontSize(self):
        self.table_widget.decreaseFontSize()

    def resetWorksheetFontSize(self):
        self.table_widget.resetFontSize()

    def SelectTheme(self):
        items = QStyleFactory.keys()
        if len(items) <= 1:
            return

        items.sort()
        item, ok = QInputDialog.getItem(self, "Select Theme", "Available Themes", items, 0, False)

        if ok:
            self.Parent.setStyle(item)
            self.currentTheme = item
            optlist = [self.currentTheme]
            with open('LaTeXTableCreatorOptions.opt', 'wb') as f:
                try:
                    pickle.dump(optlist, f)
                except:
                    QMessageBox.warning(self, "File Not Saved", "The options file could not be saved.",
                                        QMessageBox.Ok)
            # self.saveOptions()

    def createDockWidget(self):
        """
        Create dock widget which will hold the options panes.
        """
        # Set up dock widget
        self.dock_widget = QDockWidget()
        self.dock_widget.setWindowTitle("LaTeX Options")

        # Set dock attributes
        self.dock_widget.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)

        self.dock_widget.setFeatures(QDockWidget.DockWidgetFloatable |
                                     QDockWidget.DockWidgetMovable)
        self.options_pane = LTCOptionsEditorPane()

        # Create a scroll area for the options panels.
        scroll_area = QScrollArea()
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(self.options_pane)

        # Add and place the dock.
        self.dock_widget.setWidget(scroll_area)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock_widget)

    def aboutDialog(self):
        """
        Display information about program dialog box
        """
        QMessageBox.about(self, self.program_title + "  Version " + self.version,
                          self.authors + "\nVersion " + self.version +
                          "\nCopyright " + self.copyright +
                          "\nDeveloped in Python using the PySide6 GUI toolset.\n" + self.licence)

    def onHelp(self):
        """
        Display information about program dialog box
        """
        self.url_home_string = "file://" + self.resource_path("Help/index.html")
        webbrowser.open(self.url_home_string)

    def openFile(self):
        """
        Open a binary file containing a table item data and load it into the table.
        """
        file_name, _ = QFileDialog.getOpenFileName(self, "Open File",
                                                   "", "Table Data Files (*.dat);;All Files (*.*)")

        if file_name:
            with open(file_name, 'rb') as f:
                try:
                    items = pickle.load(f)
                    self.table_widget.setRowCount(1)
                    self.table_widget.setColumnCount(1)
                    self.table_widget.setCurrentCell(0, 0)
                    self.table_widget.paste(items)
                    self.setSizeSpinnersToTableSize()
                except:
                    QMessageBox.warning(self, "File Not Loaded", "The file " + file_name + " could not be loaded.",
                                        QMessageBox.Ok)

    def saveFile(self):
        """
        Save the table to a binary file.
        """
        dialog = QFileDialog()
        dialog.setFilter(dialog.filter() | QDir.Hidden)
        dialog.setDefaultSuffix('dat')
        dialog.setAcceptMode(QFileDialog.AcceptSave)
        dialog.setNameFilters(['Table Data Files (*.dat)'])
        dialog.setWindowTitle('Save As')

        if dialog.exec() == QDialog.Accepted:
            filelist = dialog.selectedFiles()
            if len(filelist) > 0:
                file_name = filelist[0]
                items = self.table_widget.getTableContents()
                with open(file_name, 'wb') as f:
                    try:
                        pickle.dump(items, f)
                    except:
                        QMessageBox.warning(self, "File Not Saved", "The file " + file_name + " could not be saved.",
                                            QMessageBox.Ok)

    def createHeaderLine(self, item, align, bold, italic, underline, divLeft, divRight):
        """
        Creates the LaTeX around an item if it is in the row or column header
        area.  Used in conjunction with longtable and tabular environments.
        """
        itemcode = item

        if underline:
            itemcode = '\\underline{' + itemcode + '}'

        if italic:
            itemcode = '\\textit{' + itemcode + '}'

        if bold:
            itemcode = '\\textbf{' + itemcode + '}'

        itemalign = align
        if divLeft:
            itemalign = '|' + itemalign

        if divRight:
            itemalign = itemalign + '|'

        itemcode = '\\multicolumn{1}{' + itemalign + '}{' + itemcode + '}'

        return itemcode

    def createLongtable(self, textable, options):
        """
        Creates the LaTeX code for either a longtable and tabular environment given
        the options.
        """
        texCode = ''

        # Get all the options from the options dictionary.
        rows = len(textable)
        cols = len(textable[0])

        border = options['Table Border']
        firstrow = options['Table Division First Row']
        allrows = options['Table Division All Rows']
        firstcol = options['Table Division First Column']
        allcols = options['Table Division All Columns']
        align = options['Table Column Align'][0].lower()

        includeColumnHeader = options['Table Column Header']
        columnHeaderAlign = options['Table Column Header Align'][0].lower()
        columnHeaderBold = options['Table Column Header Bold']
        columnHeaderItalic = options['Table Column Header Italic']
        columnHeaderUnderline = options['Table Column Header Underline']
        columnHeaderRows = options['Table Column Header Rows']

        includeRowHeader = options['Table Row Header']
        rowHeaderAlign = options['Table Row Header Align'][0].lower()
        rowHeaderBold = options['Table Row Header Bold']
        rowHeaderItalic = options['Table Row Header Italic']
        rowHeaderUnderline = options['Table Row Header Underline']
        rowHeaderColumns = options['Table Row Header Columns']

        gridtype = options['Grid Type']
        mathMode = options['Math Mode']
        stretch = options['Array Stretch']

        # Process package hint.
        if (gridtype == 'longtable'):
            texCode += '% Package: \\usepackage{longtable} \n\n'

        # Add arraystretch if selected.
        if stretch:
            texCode += '{ \n'
            texCode += '\\renewcommand{\\arraystretch}{1.0}\n\n'

        # Start table code.
        if (gridtype == 'longtable'):
            texCode += '\\begin{longtable}[l]{'
        else:
            texCode += '\\begin{tabular}{'

        # Add in alignment code.
        if border or allcols:
            texCode += '|'

        for i in range(cols):
            texCode += align
            if (i == 0) and (firstcol or allcols):
                texCode += '|'
            elif (i > 0) and allcols:
                texCode += '|'
            elif (i == cols - 1) and (border or allcols):
                texCode += '|'

        texCode += '}'

        if border or allrows:
            texCode += ' \\hline '

        texCode += '\n'

        # Load table contents.
        for i in range(rows):
            for j in range(cols):
                if mathMode:
                    itemCode = '$' + textable[i][j] + '$'
                else:
                    itemCode = textable[i][j]

                divLeft = False
                divRight = False
                if (j == 0) and (border or allcols):
                    divLeft = True

                if (j == 0) and (firstcol or allcols):
                    divRight = True
                elif (j > 0) and allcols:
                    divRight = True
                elif (j == cols - 1) and (border or allcols):
                    divRight = True

                if includeColumnHeader and (i < columnHeaderRows):
                    itemCode = self.createHeaderLine(itemCode, columnHeaderAlign, columnHeaderBold,
                                                     columnHeaderItalic, columnHeaderUnderline,
                                                     divLeft, divRight)
                elif includeRowHeader and (j < rowHeaderColumns):
                    itemCode = self.createHeaderLine(itemCode, rowHeaderAlign, rowHeaderBold,
                                                     rowHeaderItalic, rowHeaderUnderline,
                                                     divLeft, divRight)

                texCode += itemCode

                if j < cols - 1:
                    texCode += ' & '
                else:
                    texCode += ' \\\\ '

            if (i == 0) and (firstrow or allrows):
                texCode += ' \\hline '
            elif (i == rows - 1) and border:
                texCode += ' \\hline '
            elif allrows:
                texCode += ' \\hline '

            texCode += '\n'

            if columnHeaderRows > rows:
                columnHeaderRows = rows

            if (gridtype == 'longtable'):
                if includeColumnHeader and (i == columnHeaderRows - 1):
                    texCode += '\\endhead \n'
                    texCode += '\\endfoot \n'
                    texCode += '\\endlastfoot \n'

        texCode += '\\end{' + gridtype + '} \n'

        # Close stretch.
        if stretch:
            texCode += '} \n'

        return texCode

    def createTabbing(self, textable, options):
        """
        Creates the LaTeX code for the tabbing environment given the options.
        """
        texCode = ''

        rows = len(textable)
        cols = len(textable[0])

        width = options['Tabbing Column Width']
        mathMode = options['Math Mode']

        texCode += '\\begin{tabbing} \n'

        # Set column widths.
        for i in range(cols):
            texCode += '\\hspace{' + str(width) + 'pt}\\='

        texCode += '\\kill \n'

        # Load table contents.
        for i in range(rows):
            for j in range(cols):
                if mathMode:
                    texCode += '$' + textable[i][j] + '$'
                else:
                    texCode += textable[i][j]

                if j < cols - 1:
                    texCode += ' \\> '
                else:
                    texCode += ' \\\\ \n'

        texCode += '\\end{tabbing} \n'

        return texCode

    def createArray(self, textable, options):
        """
        Creates the LaTeX code for the array environment given the options.
        """
        texCode = ''

        # Get options fro options dictionary.
        rows = len(textable)
        cols = len(textable[0])

        border = options['Array Border']
        firstrow = options['Array Division First Row']
        allrows = options['Array Division All Rows']
        firstcol = options['Array Division First Column']
        allcols = options['Array Division All Columns']
        align = options['Array Column Align'][0].lower()
        mathMode = options['Math Mode']
        stretch = options['Array Stretch']

        dectype = options['Array Decoration']
        decorationLeft = ''
        decorationRight = ''
        if dectype != 'None':
            decorationLeft = dectype[0]
            decorationRight = dectype[1]

        # Process arraystretch
        if stretch:
            texCode += '{ \n'
            texCode += '\\renewcommand{\\arraystretch}{1.0}\n\n'

        if mathMode:
            texCode += '\\[ \n'

        if dectype != 'None':
            texCode += '\\left' + decorationLeft + '\n'

        # Begin Array
        texCode += '\\begin{array}{'

        # Load alignment options
        if border or allcols:
            texCode += '|'

        for i in range(cols):
            texCode += align
            if (i == 0) and (firstcol or allcols):
                texCode += '|'
            elif (i > 0) and allcols:
                texCode += '|'
            elif (i == cols - 1) and (border or allcols):
                texCode += '|'

        texCode += '}'

        if border or allrows:
            texCode += ' \\hline '

        texCode += '\n'

        # Load table contents.
        for i in range(rows):
            for j in range(cols):
                texCode += textable[i][j]

                if j < cols - 1:
                    texCode += ' & '
                else:
                    texCode += ' \\\\ '

            if (i == 0) and (firstrow or allrows):
                texCode += ' \\hline '
            elif (i == rows - 1) and border:
                texCode += ' \\hline '
            elif allrows:
                texCode += ' \\hline '

            texCode += '\n'

        texCode += '\\end{array} \n'

        if dectype != 'None':
            texCode += '\\right' + decorationRight + '\n'

        # Close math mode and stretch.
        if mathMode:
            texCode += '\\] \n'

        if stretch:
            texCode += '} \n'

        return texCode

    def createMatrix(self, textable, options):
        """
        Creates the LaTeX code for the matrix environment given the options.
        """
        texCode = ''

        # Get options
        rows = len(textable)
        cols = len(textable[0])

        mathMode = options['Math Mode']
        stretch = options['Array Stretch']

        dectype = options['Matrix Decoration']
        decorationLeft = ''
        decorationRight = ''
        if dectype != 'None':
            decorationLeft = dectype[0]
            decorationRight = dectype[1]

        # Include package hint.
        texCode += '% Package: \\usepackage{amsmath} \n\n'

        # Process arraystretch and math mode.
        if stretch:
            texCode += '{ \n'
            texCode += '\\renewcommand{\\arraystretch}{1.0}\n\n'

        if mathMode:
            texCode += '\\[ \n'

        if dectype != 'None':
            texCode += '\\left' + decorationLeft + '\n'

        texCode += '\\begin{matrix} \n'

        # Load table contents.
        for i in range(rows):
            for j in range(cols):
                texCode += textable[i][j]

                if j < cols - 1:
                    texCode += ' & '
                else:
                    texCode += ' \\\\ '

            texCode += '\n'

        texCode += '\\end{matrix} \n'

        if dectype != 'None':
            texCode += '\\right' + decorationRight + '\n'

        # Close math mode and stretch.
        if mathMode:
            texCode += '\\] \n'

        if stretch:
            texCode += '} \n'

        return texCode

    def createSpecialMatrix(self, textable, options):
        """
        Creates the LaTeX code for sprcial matrix types.
        """
        texCode = ''

        rows = len(textable)
        cols = len(textable[0])

        mathMode = options['Math Mode']
        stretch = options['Array Stretch']

        mattype = options['Special Matrix Decoration']

        # Include package hint.
        texCode += '% Package: \\usepackage{amsmath} \n\n'

        # Processs arraystretch and math mode.
        if stretch:
            texCode += '{ \n'
            texCode += '\\renewcommand{\\arraystretch}{1.0}\n\n'

        if mathMode:
            texCode += '\\[ \n'

        texCode += '\\begin{' + mattype + '} \n'

        # Load matrix contents.
        for i in range(rows):
            for j in range(cols):
                texCode += textable[i][j]

                if j < cols - 1:
                    texCode += ' & '
                else:
                    texCode += ' \\\\ '

            texCode += '\n'

        texCode += '\\end{' + mattype + '} \n'

        # Close arraystretch and math mode.
        if mathMode:
            texCode += '\\] \n'

        if stretch:
            texCode += '} \n'

        return texCode

    def createLaTeXCode(self):
        """
        Entry point for the LaTeX code creation code.  Farms out the code
        by the type of structure tht is requested.
        """
        currentTable = self.table_widget.getTableContents()
        options = self.options_pane.getOptionsInfo()

        texCode = ''
        gridtype = options['Grid Type']
        if (gridtype == 'longtable') or (gridtype == 'tabular'):
            texCode = self.createLongtable(currentTable, options)
        elif gridtype == 'tabbing':
            texCode = self.createTabbing(currentTable, options)
        elif gridtype == 'array':
            texCode = self.createArray(currentTable, options)
        elif gridtype == 'matrix':
            texCode = self.createMatrix(currentTable, options)
        elif gridtype == 'Special Matrix':
            texCode = self.createSpecialMatrix(currentTable, options)
        else:
            print('error')

        return texCode

    def latexCopy(self):
        """
        Calls the code creator and sends it to the clipboard.
        """
        texCode = self.createLaTeXCode()
        self.clipboard.setText(texCode)

    def copyAll(self):
        """
        Copies the table to the clipboard, tab delimited separation.
        """
        items = self.table_widget.getTableContents()
        str = self.itemsToTabString(items)
        self.clipboard.setText(str)

    def copySelected(self):
        """
        Copies the table selection to the clipboard, tab delimited separation.
        """
        items = self.table_widget.getSelectedTableContents()
        str = self.itemsToTabString(items)
        self.clipboard.setText(str)

    def copyMaxima(self):
        """
        Copies the table as Maxima matrix code to the clipboard.
        """
        items = self.table_widget.getTableContents()

        retstr = 'matrix('
        for i in range(len(items)):
            row = items[i]
            retstr = retstr + '['
            for j in range(len(row)):
                retstr = retstr + row[j]
                if (j < len(row) - 1):
                    retstr = retstr + ','

            retstr = retstr + ']'

            if (i < len(items) - 1):
                retstr = retstr + ','

        retstr = retstr + ')'
        self.clipboard.setText(retstr)

    def copySage(self):
        """
        Copies the table as SageMath matrix code to the clipboard.
        """
        items = self.table_widget.getTableContents()

        retstr = 'matrix(QQ, ['
        for i in range(len(items)):
            row = items[i]
            retstr = retstr + '['
            for j in range(len(row)):
                retstr = retstr + row[j]
                if (j < len(row) - 1):
                    retstr = retstr + ','

            retstr = retstr + ']'

            if (i < len(items) - 1):
                retstr = retstr + ','

        retstr = retstr + '])'
        self.clipboard.setText(retstr)

    def copyHTML(self):
        """
        Copies the table as HTML code to the clipboard.
        """
        items = self.table_widget.getTableContents()

        retstr = '<TABLE BORDER=1 CELLPADDING=1 CELLSPACING=0>\n'
        for i in range(len(items)):
            row = items[i]
            retstr = retstr + '<TR>\n'
            for j in range(len(row)):
                retstr = retstr + '<TD>' + row[j] + '</TD>'

            retstr = retstr + '\n</TR>\n'

        retstr = retstr + '</TABLE>'
        self.clipboard.setText(retstr)

    def copySpecial(self, ld, rd):
        """
        Copies the table using the specified delimiters (ld and rd) to the clipboard.
        """
        items = self.table_widget.getTableContents()
        str = self.itemsToDelimitedString(items, ld, rd)
        self.clipboard.setText(str)

    def copyGeoGebra(self):
        """
        Copies the table as GeoGebra code {} to the clipboard.
        """
        self.copySpecial('{', '}')

    def copyBracket(self):
        """
        Copies the table [] delimited to the clipboard.
        """
        self.copySpecial('[', ']')

    def copyAngleBracket(self):
        """
        Copies the table <> delimited to the clipboard.
        """
        self.copySpecial('<', '>')

    def selectall(self):
        """
        Selects the entire table.
        """
        self.table_widget.selectAll()

    def paste(self):
        """
        Pastes the clipboard contents, assumed to be tab delimited, to the table.
        """
        str = self.clipboard.text()
        items = self.tabStringToItems(str)
        self.table_widget.paste(items)
        self.setSizeSpinnersToTableSize()

    def pasteLatex(self):
        """
        Pastes the clipboard contents, assumed to be LaTeX format, to the table.
        """
        items = []
        str = self.clipboard.text()
        str = str.replace('\\&', '<~~amp~~>')
        str = str.replace('\\hline', '')
        strlist = str.split('\\\\')

        for line in strlist:
            line = line.rstrip().lstrip()
            if line != '':
                linelist = line.split('&')
                for i in range(len(linelist)):
                    newline = linelist[i]
                    newline = newline.replace('<~~amp~~>', '\\&')
                    newline = newline.rstrip().lstrip()
                    linelist[i] = newline
                items.append(linelist)

        self.table_widget.setRowCount(1)
        self.table_widget.setColumnCount(1)
        self.table_widget.setCurrentCell(0, 0)
        self.table_widget.paste(items)
        self.setSizeSpinnersToTableSize()

    def setSizeSpinnersToTableSize(self):
        """
        Resets the spinner values to match the size of the table.
        """
        self.rows.blockSignals(True)
        self.columns.blockSignals(True)
        if self.table_widget.rowCount() != self.rows.value():
            self.rows.setValue(self.table_widget.rowCount())
        if self.table_widget.columnCount() != self.columns.value():
            self.columns.setValue(self.table_widget.columnCount())
        self.rows.blockSignals(False)
        self.columns.blockSignals(False)

    def itemsToDelimitedString(self, items, ld, rd):
        """
        Converts a list of row lists of table elements to a string delimited by ld and rd
        characters.
        """
        retstr = ld
        for i in range(len(items)):
            row = items[i]
            retstr = retstr + ld
            for j in range(len(row)):
                retstr = retstr + row[j]
                if (j < len(row) - 1):
                    retstr = retstr + ','

            retstr = retstr + rd

            if (i < len(items) - 1):
                retstr = retstr + ','

        retstr = retstr + rd
        return retstr

    def itemsToTabString(self, items):
        """
        Converts a list of row lists of table elements to a tab delimited string.
        """
        retstr = ''
        for i in range(len(items)):
            row = items[i]
            for j in range(len(row)):
                retstr = retstr + row[j]
                if (j == len(row) - 1):
                    retstr = retstr + '\n'
                else:
                    retstr = retstr + '\t'
        return retstr

    def tabStringToItems(self, tdstr):
        """
        Converts a tab delimited string to a list of row lists of table elements.
        """
        retitems = []
        strobj = str(tdstr)
        strlist = strobj.split('\n')
        for line in strlist:
            if len(line) > 0:
                splitline = line.split('\t')
                retitems.append(splitline)
        return retitems

    def addRowAbove(self):
        """
        Adds a row above the current selected row.
        """
        self.table_widget.addRowAbove()
        self.setSizeSpinnersToTableSize()

    def addRowBelow(self):
        """
        Adds a row below the current selected row.
        """
        self.table_widget.addRowBelow()
        self.setSizeSpinnersToTableSize()

    def addColumnBefore(self):
        """
        Adds a column before the current selected column.
        """
        self.table_widget.addColumnBefore()
        self.setSizeSpinnersToTableSize()

    def addColumnAfter(self):
        """
        Adds a column after the current selected column.
        """
        self.table_widget.addColumnAfter()
        self.setSizeSpinnersToTableSize()

    def deleteRows(self):
        """
        Delete the selected rows.
        """
        self.table_widget.deleteRows()
        self.setSizeSpinnersToTableSize()

    def deleteColumns(self):
        """
        Delete the selected columns.
        """
        self.table_widget.deleteColumns()
        self.setSizeSpinnersToTableSize()

    def deleteRowsColumns(self):
        """
        Delete the selected rows and columns.
        """
        self.table_widget.deleteRowsColumns()
        self.setSizeSpinnersToTableSize()

    def clearTable(self):
        """
        Clears the entire table.
        """
        self.table_widget.clearTable()

    def newtable(self):
        """
        Creates a new instance of the program and launches it.  The
        programList variable keeps the new instance from losing its
        link and being automatically removed.
        """
        # Seems to work, look for "better" solution.
        self.newwindow = LaTeXTableEditor()
        self.programList.append(self.newwindow)

    def transpose(self):
        """
        Transposes the table.
        """
        self.table_widget.transpose()
        self.setSizeSpinnersToTableSize()

    def trim(self):
        """
        Trims whitespace from the front and back of all table elements.
        """
        self.table_widget.trimcells()

    def filltext(self):
        """
        Gets a string from the user and fills the table with that string.
        """
        fill_text, ok = QInputDialog.getText(self, "Fill Cells With", "Text:")
        if ok:
            self.table_widget.fillcells(fill_text)

    def adjustWidths(self):
        """
        Set the column widths to fit the size of the data.
        """
        self.table_widget.resizeColumnsToContents()

    def adjustHeights(self):
        """
        Set the row heights to fit the size of the data.
        """
        self.table_widget.resizeRowsToContents()

    def adjustWidthHeight(self):
        """
        Set the row heights and column widths to fit the size of the data.
        """
        self.table_widget.resizeColumnsToContents()
        self.table_widget.resizeRowsToContents()

    def undo(self):
        """
        Undo the last edit.
        """
        self.table_widget.undo()
        self.setSizeSpinnersToTableSize()

    def redo(self):
        """
        Redo the last edit.
        """
        self.table_widget.redo()
        self.setSizeSpinnersToTableSize()


if __name__ == '__main__':
    """
    Initiate the program. 
    """
    app = QApplication(sys.argv)
    window = LaTeXTableEditor(app)
    progcss = LTCappcss()
    app.setStyleSheet(progcss.getCSS())
    sys.exit(app.exec())
