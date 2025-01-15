#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Aug  6 11:57:31 2021
Revised: 8/10/2021

@author: Don Spickler

This program allows the user to input a table into a standard spreadsheet
like grid and then export the LaTeX code to the clipboard given the options
the user selects.

"""

import pickle
import sys
import os

from PySide6.QtCore import Qt, QSize, QDir
from PySide6.QtGui import QIcon, QAction
from PySide6.QtWidgets import (QApplication, QMainWindow, QStatusBar,
                               QToolBar, QDockWidget, QSpinBox, QHBoxLayout,
                               QVBoxLayout, QWidget, QLabel, QScrollArea, QMessageBox,
                               QInputDialog, QFileDialog, QDialog)

from CSS_Class import appcss
from LC_Table import LC_Table
from OptionsPane import OptionsEditorPane

import webbrowser

# For the Mac OS
os.environ['QT_MAC_WANTS_LAYER'] = '1'


class LaTeXTableEditor(QMainWindow):

    def __init__(self, parent=None):
        """
        Initialize the program and set up the programList structure that
        allows the creation of child applications.
        """
        super().__init__()
        self.Parent = parent
        
        self.authors = "Don Spickler"
        self.version = "2.4.1"
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
        icon = QIcon(self.resource_path("icons/ProgramIcon.png"))
        self.setWindowIcon(icon)

        self.createTablePane()
        self.createMenu()
        self.createToolBar()
        self.createDockWidget()
        self.show()

    def resource_path(self, relative_path):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)

    def createTablePane(self):
        """
        Sets up the portion of the screen that contains the table and the
        size selection spinners.
        """
        self.table_widget = LC_Table()

        row_label = QLabel("Rows")
        self.rows = QSpinBox()
        self.rows.setRange(1, 10000)
        self.rows.setMinimumWidth(75)
        self.rows.setValue(3)
        self.rows.editingFinished.connect(self.resizeTable)

        column_label = QLabel("Columns")
        self.columns = QSpinBox()
        self.columns.setRange(1, 1000)
        self.columns.setMinimumWidth(75)
        self.columns.setValue(3)
        self.columns.editingFinished.connect(self.resizeTable)

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

        self.adjust_widths_act = QAction("Adjust Column Widths", self)
        self.adjust_widths_act.setStatusTip('Adjust the column widths to fit the contents.')
        self.adjust_widths_act.triggered.connect(self.adjustWidths)

        self.adjust_heights_act = QAction("Adjust Row Heights", self)
        self.adjust_heights_act.setStatusTip('Adjust the row heights to fit the contents.')
        self.adjust_heights_act.triggered.connect(self.adjustHeights)

        self.adjust_width_height_act = QAction("Adjust Row and Column Sizes", self)
        self.adjust_width_height_act.setShortcut('Ctrl+R')
        self.adjust_width_height_act.setStatusTip('Adjust the row and column sizes to fit the contents.')
        self.adjust_width_height_act.triggered.connect(self.adjustWidthHeight)

        # Create edit menu actions
        self.copy_selected_act = QAction(QIcon(self.resource_path('icons/copy.png')), "Copy Selected", self)
        self.copy_selected_act.setShortcut('Ctrl+C')
        self.copy_selected_act.setStatusTip('Copy table selection to the clipboard.')
        self.copy_selected_act.triggered.connect(self.copySelected)

        self.copy_all_act = QAction("Copy All", self)
        self.copy_all_act.setShortcut('Shift+Ctrl+C')
        self.copy_all_act.setStatusTip('Copy entire table to the clipboard.')
        self.copy_all_act.triggered.connect(self.copyAll)

        self.paste_act = QAction(QIcon(self.resource_path('icons/paste.png')), "Paste", self)
        self.paste_act.setShortcut('Ctrl+V')
        self.paste_act.setStatusTip('Paste clipboard to table.')
        self.paste_act.triggered.connect(self.paste)

        self.copy_latex_act = QAction(QIcon(self.resource_path('icons/copylatex.png')), "Copy as LaTeX", self)
        self.copy_latex_act.setShortcut('Ctrl+L')
        self.copy_latex_act.setStatusTip('Copy table to LaTeX code with selected options.')
        self.copy_latex_act.triggered.connect(self.latexCopy)

        self.paste_from_latex_act = QAction("Paste from LaTeX", self)
        self.paste_from_latex_act.setShortcut('Shift+Ctrl+L')
        self.paste_from_latex_act.setStatusTip('Paste from LaTeX code to table.')
        self.paste_from_latex_act.triggered.connect(self.pasteLatex)

        self.copy_maxima_act = QAction("Copy as Maxima", self)
        self.copy_maxima_act.setStatusTip('Copy matrix to Maxima code.')
        self.copy_maxima_act.triggered.connect(self.copyMaxima)

        self.copy_sage_act = QAction("Copy as SageMath", self)
        self.copy_sage_act.setStatusTip('Copy matrix to SageMath code.')
        self.copy_sage_act.triggered.connect(self.copySage)

        self.copy_mathematica_act = QAction("Copy as Mathematica", self)
        self.copy_mathematica_act.setStatusTip('Copy matrix to Mathematica code.')
        self.copy_mathematica_act.triggered.connect(self.copyMathematica)

        self.copy_bracket_act = QAction("Copy [...] Delimited", self)
        self.copy_bracket_act.setStatusTip('Copy matrix in [...] delimited form.')
        self.copy_bracket_act.triggered.connect(self.copyBracket)

        self.copy_curleybracket_act = QAction("Copy {...} Delimited", self)
        self.copy_curleybracket_act.setStatusTip('Copy matrix in {...} delimited form.')
        self.copy_curleybracket_act.triggered.connect(self.copyMathematica)

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

        # Create help menu actions
        self.help_about_act = QAction(QIcon(self.resource_path('icons/About.png')), "About...", self)
        self.help_about_act.setStatusTip('About the LaTeX Table Creator')
        self.help_about_act.triggered.connect(self.aboutDialog)

        self.help_help_act = QAction("Help...", self)
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
        edit_menu.addAction(self.paste_act)
        edit_menu.addSeparator()
        edit_menu.addAction(self.copy_latex_act)
        edit_menu.addAction(self.paste_from_latex_act)
        edit_menu.addSeparator()
        edit_menu.addAction(self.copy_sage_act)
        edit_menu.addAction(self.copy_maxima_act)
        edit_menu.addAction(self.copy_mathematica_act)
        edit_menu.addSeparator()
        edit_menu.addAction(self.copy_bracket_act)
        edit_menu.addAction(self.copy_curleybracket_act)
        edit_menu.addAction(self.copy_angle_bracket_act)
        edit_menu.addAction(self.copy_html_act)
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
        table_menu.addAction(self.select_all_act)
        table_menu.addAction(self.clear_table_act)
        table_menu.addSeparator()
        table_menu.addAction(self.transpose_act)
        table_menu.addAction(self.trim_act)
        table_menu.addAction(self.fill_cells_act)
        table_menu.addSeparator()
        table_menu.addAction(self.adjust_widths_act)
        table_menu.addAction(self.adjust_heights_act)
        table_menu.addAction(self.adjust_width_height_act)

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
        tool_bar.setIconSize(QSize(16, 16))
        self.addToolBar(tool_bar)

        # Add actions to toolbar
        tool_bar.addAction(self.file_new_act)
        tool_bar.addAction(self.file_open_act)
        tool_bar.addAction(self.file_saveas_act)
        tool_bar.addSeparator()
        tool_bar.addAction(self.copy_selected_act)
        tool_bar.addAction(self.paste_act)
        tool_bar.addSeparator()
        tool_bar.addAction(self.Undo_act)
        tool_bar.addAction(self.Redo_act)
        tool_bar.addSeparator()
        tool_bar.addAction(self.copy_latex_act)
        tool_bar.addSeparator()
        tool_bar.addAction(self.help_about_act)

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
        self.options_pane = OptionsEditorPane()

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
        self.url_home_string = "file://" + self.resource_path("Help/Help.html")
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

    def copyMathematica(self):
        """
        Copies the table as Mathematica code {} to the clipboard.
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
        str = str.replace('\&', '<~~amp~~>')
        str = str.replace('\hline', '')
        strlist = str.split('\\\\')

        for line in strlist:
            line = line.rstrip().lstrip()
            if line != '':
                linelist = line.split('&')
                for i in range(len(linelist)):
                    newline = linelist[i]
                    newline = newline.replace('<~~amp~~>', '\&')
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
        if self.table_widget.rowCount() != self.rows.value():
            self.rows.setValue(self.table_widget.rowCount())
        if self.table_widget.columnCount() != self.columns.value():
            self.columns.setValue(self.table_widget.columnCount())

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
    progcss = appcss()
    app.setStyleSheet(progcss.getCSS())
    sys.exit(app.exec())
