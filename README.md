# XML-Notepad `latest ver 2.6`

XML-Notepad is a Python-based desktop application built with Tkinter for viewing, navigating, and editing XML files, with a special focus on identifying and working with tabular data within XML structures. It allows users to open large XML files, view their hierarchical structure, inspect node details, and automatically detect and display repetitive child elements as editable tables. These tables can then be modified, sorted, and exported to CSV.

![XML-Notepad Screenshot](https://github.com/user-attachments/assets/80434d3b-d66d-433a-b9b5-f2aadda7c72d)

## Key Features

### XML Viewing & Navigation
- Open and parse large XML files efficiently (using threading and chunking).
- Display XML structure in a hierarchical tree view.
- View details of selected XML nodes (tag, text, attributes).
- Context menu on XML tree nodes for actions like "Export Node as CSV".

### Automatic Table Detection & Editing (UNIQUE FEATURE)
- Intelligently identifies and extracts tabular data from XML structures based on repetitive child tags.
- Displays detected tables in a user-friendly, editable grid (Treeview).
- Select different detected tables via a combobox.
- In-cell editing of table data (double-click).
- Row operations: Cut, Copy, Paste, Delete specific to the current table.
- Column sorting by clicking headers.
- **Column resizing**:
  - Auto-adjust column widths to content (double-click header or Edit menu).
  - Manual resize from Edit menu (fits to content within limits).
- Find functionality within tables (search specific fields for values).

### Advanced Query Designer (NEW FEATURE)
Access via `Utils > Query Designer...`.

#### Visual Designer Tab
- Select two tables (Table 1 and Table 2) from the loaded XML.
- Define join conditions between fields from Table 1 and Table 2. Supports single or compound keys.
- Select desired output fields (columns) from both tables using a shuttle control.
- Reorder output fields using "Up" and "Down" buttons to define report column order.

![image](https://github.com/user-attachments/assets/d35d8b4f-1b57-4dd0-b731-9be858c63894) 
Main Visual designer  


#### Query View Tab
- View a SQL-like representation of the visually designed query (auto-updates).
- Toggle Edit Mode to manually edit the query text directly.
- **Validate & Apply to Designer**: Parse manually entered query text and update the visual designer's state if valid.

![image](https://github.com/user-attachments/assets/0d763d97-f721-4c8f-9d89-ef45fbfb9252) 
Manually query editor

#### Query Options
- Choose between Inner Join (matching records) and Left Anti-Join (records from Table 1 not matching Table 2).
- Optionally limit the number of results (e.g., Top 50, 100, or custom value).

#### Results Display
- Display query results in a dedicated grid (Treeview).
- Includes a "Match_Count" column for Inner Joins, showing the number of Table 2 records linked to each Table 1 record.
- Sort results by clicking on any column header (toggles ascending/descending order, shows arrow indicator).
- Export results to CSV with a dedicated "Export as CSV..." button.

#### Save/Load Query
- Save the entire query configuration (tables, joins, output fields, options) as a JSON file via `Query > Save Query...`.
- Load a previously saved query configuration via `Query > Load Query...`, re-populating the designer.

### Data Export & Import
- Save modified XML data to a new file (`Save XML As...`).
- Export the currently displayed table to a CSV file (`Save Table as CSV...`, `Ctrl+S`).
- Export a selected XML node and its uniform children to a CSV file.

### User Experience
- Undo/Redo functionality for table row operations (Cut, Paste, Delete) and cell edits.
- Status bar with progress display for long operations (loading, saving).
- Keyboard shortcuts for common actions (Open: `Ctrl+O`, Save Table: `Ctrl+S`, Find: `Ctrl+F`).
- Responsive UI with paned windows for adjustable layout.
- Error handling and informative messages.

## Requirements
- Python 3.6+
- Tkinter (usually included with standard Python installations)
- No external libraries are required.

## How to Run
1. Save the code as a Python file (e.g., `xml_notepad.py`).
2. Open a terminal or command prompt.
3. Navigate to the directory where you saved the file.
4. Run the script using:
   ```bash
   python xml_notepad.py
   ```
# Usage Guide

### 1. Opening an XML File
- Go to `File > Open XML...` (or press `Ctrl+O`).
- Select your XML file.
- A progress bar will indicate loading status for larger files.
- The XML structure will appear in the "XML Tree" panel on the left.
- The loaded filename will be displayed at the top.

### 2. Navigating the XML Tree
- Click on nodes in the "XML Tree" panel to expand or collapse them.
- When a node is selected:
  - Its details (tag, text, attributes, or child summary) will appear in the "Node Details" panel on the right.
  - If the node is a parent of uniformly tagged children suitable for CSV export, the "Export Node as CSV..." option in the right-click context menu will be enabled.

### 3. Working with Detected Tables
**Table Detection**: The application automatically scans the loaded XML for parent elements containing multiple child elements with the same tag (configurable by `MIN_ROWS_FOR_TABLE` and `MIN_PERCENT_SIMILAR`).

#### Selecting a Table
- If tables are detected, they will be listed in the "Tables:" combobox at the top.
- Select a table name from the combobox.
- The right panel will switch to "Table View", displaying the data in a grid.
- The table title will indicate the selected table and the number of rows.

#### Viewing Table Data
- The table shows a row number (#) and columns derived from the child elements' tags.
- Use scrollbars to navigate large tables.

#### Editing Table Cells
- Double-click on a cell (except the # column) to edit its content.
- An entry widget will appear. Type your changes.
- Press `Enter` to commit the change or `Escape` to cancel.
- Changes are reflected in the underlying XML data.

#### Row Operations (Edit Menu)
- First, select a row by clicking its row number (# column) or any cell within it (the entire row will highlight).
- **Cut Row**: Copies the row to an internal clipboard and removes it from the table and XML.
- **Copy Row**: Copies the row to the internal clipboard.
- **Paste Row**: Pastes the clipboard row data as a new row at the end of the current table (only if pasting into the same table type).
- **Delete Row**: Removes the selected row from the table and XML.

#### Sorting Table Data
- Click on a column header to sort the table by that column.
- Clicking the same header again toggles between ascending and descending order.

#### Resizing Columns
- `Edit > Resize Columns`: Adjusts all column widths to fit their content (up to a maximum width).
- Double-click any column header: Auto-adjusts all column widths to fit content (stretches to fill available space).

#### Finding Data in Tables
- Go to `Edit > Find...` (or press `Ctrl+F`).
- In the "Find in Table" dialog:
  - Select the Table to search in.
  - Select the Field (column) to search.
  - Enter the Value to find.
  - Optionally, check Match case.
  - Click Find. Matching rows will be listed.
  - Clicking a row in the results or using Next/Previous will highlight and jump to that row in the main table view.

### 4. Advanced Query Designer Usage
- Go to `Utils > Query Designer...`.

#### Visual Designer Tab
- In "1. Select Tables", choose your Left (T1) and Right (T2) tables.
- In "2. Define Join Conditions", select corresponding fields from T1 and T2, then click "Add Join ->". Repeat for compound keys.
- In "3. Design Report Output", use the `>` and `<` buttons to move fields between "Available Fields" and "Selected Fields". Use "Up" and "Down" buttons to reorder columns in the output report.
- ** NEW ** ability to use only 1 table. When you use only 1 table then the you can choose filter conditions instead of JOIN conditions.

![image](https://github.com/user-attachments/assets/0087cac5-d8de-46da-b93b-e7201934b7c2)


#### Query View Tab
- Observe the dynamically generated SQL-like query.
- Click "Toggle Edit Mode" to manually modify the query.
- Typing a table name like `T1:` now shows an intellsense list of fields for that table to quickly type in the fieldss. Fields are selected by double-clicking or pressing enter/return key.
- After editing, click "Validate & Apply to Designer" to parse your changes and update the visual designer.

![image](https://github.com/user-attachments/assets/23f69294-cc46-4980-8125-e103b13715a7)


#### Set Options and Run
- In "4. Set Options and Run", select "Inner Join" for matching records or "Left Anti-Join" for T1 records with no T2 match.
- Optionally, check "Limit:" and set a value to restrict the number of returned rows.
- Click "Run Query" to execute the query and display results in the "Results" grid below.

#### Query Results
- Results will populate the grid.
- Click on a column header to sort the results. Clicking again reverses the sort order. An arrow indicator (▲/▼) will appear on the sorted column's header.
- Click "Export as CSV..." to save the query results to a CSV file.

#### Save/Load Queries (Query Menu)
- From the Query Designer window's menu bar, go to `Query > Save Query...` to save the current query configuration.
- Go to `Query > Load Query...` to load a previously saved query and populate the designer.

### 5. Saving Data
#### Save XML As...
- Go to `File > Save XML As...`.
- Choose a location and filename. This saves the entire current state of the XML tree, including any modifications made via table editing.

#### Save Table as CSV...
- Ensure a table is currently displayed in the right panel.
- Go to `File > Save Table as CSV...` (or press `Ctrl+S` when a table is active).
- Choose a location and filename to save the currently viewed table data as a CSV file.

### 6. Exporting Node as CSV
- In the "XML Tree" view (left panel), right-click on a parent node whose children you want to export.
- If the children are uniformly tagged and suitable, "Export Node as CSV..." will be enabled.
- Select it, choose a location and filename. The direct children of this node (and their sub-element text or attributes) will be exported.

### 7. Undo/Redo
- `Edit > Undo` (or `Ctrl+Z` equivalent, though not explicitly bound in this version for undo/redo)
- `Edit > Redo` (or `Ctrl+Y` equivalent)
- These actions apply to table cell edits and row operations (Cut, Paste, Delete). The undo stack has a limited size (`UNDO_STACK_SIZE`).
- Currently restricted to 20 undo/redo operations. This can be configured within the program. Do not change this setting if you have limited RAM (< 16GB).

## Configuration Constants (in code)
- `CHUNK_SIZE = 1024 * 1024`: Size of chunks (1MB) for reading large XML files.
- `MIN_ROWS_FOR_TABLE = 3`: Minimum number of similar child elements required to be considered a potential table.
- `MIN_PERCENT_SIMILAR = 0.6`: Minimum percentage of children under a parent that must share the most common tag for table detection.
- `UNDO_STACK_SIZE = 20`: Maximum number of actions stored in the undo/redo stacks.

## Limitations & Known Issues
- **Table Detection**: The heuristic for table detection (based on common child tags) might not correctly identify all desired tabular structures in complex or unusually formatted XML.
- **Performance**: While optimized for large files, extremely deep XML structures or tables with a very large number of columns/rows might still experience some UI lag during rendering or operations.
- **Undo/Redo Scope**: Undo/Redo is currently implemented for table cell edits and row-level operations (cut, copy, paste, delete) within the table view. It does not cover direct XML structural changes outside of these table operations.
- **XML Modification**: Direct editing of XML tags, attributes, or text outside the table view (e.g., in the Node Details panel or XML Tree) is not supported. Modifications are primarily through the table interface.
- **Single File Focus**: The application works with one XML file at a time.
- **Manual Query Parsing**: While improved, manual SQL-like query parsing relies on specific patterns. Complex or highly custom SQL syntax might not be parsed correctly.

## Future Enhancements (Ideas)
- Direct editing of XML attributes and text in the "Node Details" panel.
- Ability to add/delete XML nodes directly in the tree view.
- More sophisticated table detection algorithms or manual table definition.
- XPath searching capabilities.
- Customizable UI themes.
- Drag-and-drop for opening files.
- Validation against DTD/Schema.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
This project is open-source and released under the MIT License.  
(c) Shuvro Basu, 2025.
