# LaTeX Table Creator

**Download the Current Version (2.5.1)**

**Windows Users**

This is a Python application using the PySide6 GUI API, but you do not need to have either Python nor PySide6 installed on your machine to run this program. The Windows distribution of this program is as a single stand-alone executable file, LaTeXTableCreator.exe.  This software has been tested on both Windows 10 and 11.

- Download the **[LaTeXTableCreator.exe](https://github.com/mathprofdes/LaTeX-Table-Creator/releases/download/v2.5.1/LaTeXTableCreator.exe)** file.
- From Windows Explorer double-click the LaTeXTableCreator.exe file.

**Linux and MacOS Users**

This is a Python application using the PySide6 GUI API. To run this program from the source code you will need both Python3 and the python package PySide6 installed on your system.

To run this program from the source code:

- Download and extract the Source Code file from the most current release.
- Make sure that Python3 and the python package PySide6 are installed on your system.
- Run the following command from your terminal: `python LaTeXTableCreator.py` or possibly `python3 LaTeXTableCreator.py`

**Notes:** 
- For Linux and MacOS users, depending on how your system is set up, you may be able to simply double-click the LaTeXTableCreator.py file from your file browser instead of running this from the terminal.
- There are png and ico files of a program icon included if you wish to use them for a shortcut to the program. 
  - **[ProgramIcon.ico](https://github.com/mathprofdes/LaTeX-Table-Creator/releases/download/v2.5.1/ProgramIcon.ico)**
  - **[ProgramIcon.png](https://github.com/mathprofdes/LaTeX-Table-Creator/releases/download/v2.5.1/ProgramIcon.png)**

**Updates:** 
- Fixed a small bug in the Paste from LaTeX code.
- Revised the help system to Sphinx.

---

**Program Description**

The LaTeX Table Creator is an application that allows the user to input data into a grid and export the contents to LaTeX syntax given a few options. The program is not a WYSIWYG interface and certainly does not provide a complete set of options. This program is for quick conversions of table data into either table or matrix LaTeX code. It does offer copy and paste capabilities within the program and between the program and most spreadsheets (tab delimited text transfer). In addition, there are several options for populating the grid, transposing, resizing, inserting and deleting rows and columns, undo and redo, and file saving and loading of the data grid.

The LaTeX export is done through the system clipboard. The user should populate the grid with the desired data, select the LaTeX options on the right side of the window and then copy the grid as LaTeX code. From there the user can paste the code into any editor they are using to create their document.

The program currently supports longtable, tabular, tabbing, array, matrix, pmatrix, bmatrix, vmatrix, and Vmatrix environments. When copied, the clipboard text will have a commented line of any needed packages to be included in the preamble of the document. Each of the supported environments has a set of options for that environment, which includes alignment options, border and division options, header row and column creation, automatic math mode inclusion, and matrix decorations.

This program is designed to make the creation of LaTeX tables easier but is not designed to do everything for the user. For someone who is familiar with LaTeX typesetting and the basic code for tables it will provide a nice layout that should be easy to edit and manipulate. In addition, there are options for exporting the grid contents to SageMath, Maxima, and Mathematica code as well as [ ] and < > delimited strings that are commonly used in other packages.

---

**Screenshot**

![Screenshot of program.](https://github.com/mathprofdes/LaTeX-Table-Creator/releases/download/v2.5.1/LatexTableCreatorPython.png)
