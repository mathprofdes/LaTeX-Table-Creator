:html_theme.sidebar_left.remove:

===============
Program Options
===============

File Options
------------

The file options allow you to save and load table information. The saving and loading of table information is just on the current table itself, the undo and redo history is not saved.

New
^^^

The New option will open another copy of the program with a blank grid.

Open
^^^^

The Open option will open a table data file. When a file is opened it will replace the current table with the saved one. The undo/redo history is not cleared so the table that is overwritten is still in the history. The data files are stored in binary format and cannot be edited with an outside editor.

Save As...
^^^^^^^^^^

The Save As... option will save the current table to a table data file. Saving will save only the current table and not the history. The data files are stored in binary format and cannot be edited with an outside editor.

Edit Options
------------

The Edit menu contains standard copy and paste options as one would expect. It also has options for creating LaTeX code, as the program was designed to do, and export to specialized formats.


Copy Selected, Copy All, Select All, & Paste Options
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

These are your usual options for clipboard transfer to and from the grid. Copy Selected copies the currently selected portion of the table as tab delimited text to the clipboard. Copy All does the same but with the entire grid. Paste will paste in tab delimited text to the table at the currently selected position. With the paste option, if what is being pasted at the current position needs more space, rows or columns, the grid will automatically resize itself to fit the added content.

Copy as LaTeX
^^^^^^^^^^^^^

The Copy as LaTeX option will create LaTeX code for the current grid contents using the options you have selected in the LaTeX Options box. This code is automatically copied to the clipboard so you can paste it into your LaTeX document. Details about what these options do is discussed in the  :doc:`exportops` section.  For example, if you use a long table with the default options the program would put the following into the clipboard.

.. code-block:: latex

    % Package: \usepackage{longtable}

    \begin{longtable}[l]{lll}
    1 & 4 & 7 \\
    2 & 5 & 8 \\
    3 & 6 & 9 \\
    \end{longtable}

Paste from LaTeX
^^^^^^^^^^^^^^^^

There is also an option to paste from LaTeX code, Shift+Ctrl+L or using the menu option Edit > Paste from LaTeX. This will take the body of LaTeX code, extract the cell entries, and load the data into the grid. As with exporting to LaTeX code the entire grid is replaced with the data, not just the selection. For example, if you had the following table,

.. code-block:: latex

    \begin{longtable}[l]{lll}
    1 & 4 & 7 \\
    2 & 5 & 8 \\
    3 & 6 & 9 \\
    \end{longtable}

and then copied the body portion to the clipboard, not including the begin and end statements, that is,

.. code-block:: latex

    1 & 4 & 7 \\
    2 & 5 & 8 \\
    3 & 6 & 9 \\

then select Shift+Ctrl+L or Edit > Paste from LaTeX the program would extract the contents into the grid,

.. code-block:: console

    1  4  7
    2  5  8
    3  6  9

This option will also parse through \hline code but will not process multicolumn commands and some other specialized content.

Copy as SageMath
^^^^^^^^^^^^^^^^

The Copy as SageMath will formulate the grid as a SageMath matrix. For example, the standard grid

.. code-block:: console

    1  4  7
    2  5  8
    3  6  9

will copy as

.. code-block:: console

    matrix(QQ,[[1,4,7],[2,5,8],[3,6,9]])

Copy as Maxima
^^^^^^^^^^^^^^

The Copy as Maxima will formulate the grid as a Maxima matrix. For example, the standard grid

.. code-block:: console

    1  4  7
    2  5  8
    3  6  9

will copy as

.. code-block:: console

    matrix([1,4,7],[2,5,8],[3,6,9])

Copy as GeoGebra
^^^^^^^^^^^^^^^^

The Copy as GeoGebra will formulate the grid as a GeoGebra matrix, that is { } delimited. For example, the standard grid

.. code-block:: console

    1  4  7
    2  5  8
    3  6  9

copies as

.. code-block:: console

    {{1,4,7},{2,5,8},{3,6,9}}

Copy [...] Delimited
^^^^^^^^^^^^^^^^^^^^

This copies as [ ] delimited string. For example, the standard grid

.. code-block:: console

    1  4  7
    2  5  8
    3  6  9

copies as

.. code-block:: console

    [[1,4,7],[2,5,8],[3,6,9]]

Copy {...} Delimited
^^^^^^^^^^^^^^^^^^^^

This copies as { } delimited string. For example, the standard grid

.. code-block:: console

    1  4  7
    2  5  8
    3  6  9

copies as

.. code-block:: console

    {{1,4,7},{2,5,8},{3,6,9}}

Copy <...> Delimited
^^^^^^^^^^^^^^^^^^^^

This copies as < > delimited string. For example, the standard grid

.. code-block:: console

    1  4  7
    2  5  8
    3  6  9

copies as

.. code-block:: console

    <<1,4,7>,<2,5,8>,<3,6,9>>

Copy as HTML
^^^^^^^^^^^^

This copies the table to HTML markup. For example, the standard grid

.. code-block:: console

    1  4  7
    2  5  8
    3  6  9

copies as

.. image:: HTMLEx.png

Undo & Redo
^^^^^^^^^^^

The program also has undo and redo features, using the menu or the standard Ctrl+Z and Ctrl+Shift+Z respectively. Every time the grid is changed the undo history is updated with the new grid. The program does not limit the number of undos that are possible.


Table Options
-------------

The table options allow you to manipulate the data grid. Adding and deleting rows and columns, adjusting of cell sizes, and a few specialized operations.

Adding Rows and Columns
^^^^^^^^^^^^^^^^^^^^^^^

Adding rows and columns will insert a row or column either before or after the selected region. There must be at least one cell currently selected on the grid or no row or column will be inserted. If multiple cells are selected the added row will be added directly above or below the selected block of cells. The same is true for added columns.

Deleting Rows and Columns
^^^^^^^^^^^^^^^^^^^^^^^^^

When deleting rows or columns, the program will delete all of the rows or columns in the selected region. Again at least one cell must be selected for any rows or columns to be deleted.

Grid Manipulation Options
^^^^^^^^^^^^^^^^^^^^^^^^^

The Table menu has several manipulation options that come in handy from time to time.

* The Transpose option will transpose the grid, that is, turn rows to columns and columns to rows.
* The Trim option will remove any leading and trailing spaces for each data item in the grid.
* The Fill option will allow the user to input some text and it will fill all selected cells with that text.

Clearing
^^^^^^^^

The Clear All option will remove all data from the grid. Note that this does not clear the grid history. So if you mistakenly clear the grid an undo will bring the data back.

View Options
------------

Cell Size Adjustments
^^^^^^^^^^^^^^^^^^^^^

The automatic size adjustments will adjust the row height and column widths to match the data in the grid. The default size on startup and when adding rows or columns may not fit the contents of the cells, these options will make those adjustments.

Table Font Size
^^^^^^^^^^^^^^^

These options allow the user to change the size of the table font by increasing or decreasing the size by one point each time an to reset the font size to the default size.
