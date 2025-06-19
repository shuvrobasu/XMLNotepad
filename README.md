# XML/CSV-Notepad `latest ver 3`

# XML/CSV-Notepad

XML-Notepad is a Python-based desktop application built with Tkinter for viewing and navigating XML files, with a powerful focus on identifying, querying, and working with tabular data within XML structures. It allows users to open large XML files, automatically detect and display repetitive elements as tables, and perform complex data operations using an intuitive query designer.

![image](https://github.com/user-attachments/assets/f6c729e6-a986-4add-a45c-a28d12e5d271)


## Key Features

### XML Viewing & Navigation
- Open and parse large XML files efficiently.
- Display XML structure in a hierarchical tree view.
- View details of selected XML nodes (tag, text, attributes).

### Automatic Table Detection
- Intelligently identifies and extracts tabular data from XML structures.
- Displays detected tables in a user-friendly grid.

### Advanced Query Designer (NEW FEATURE)
Access via Utils > Query Designer....

#### Visual Designer Tab
**Build Complex Queries Visually:**
- Select one or two tables for querying (SELECT...FROM) or joining (JOIN).
- **Intuitive Condition Builder:** Create complex WHERE and ON clauses using a sequential builder. Add conditions, logical operators (AND, OR, NOT), and parentheses () to control the order of operations.
- **Live Validation & Syntax Highlighting:** The condition builder provides immediate feedback, highlighting operators and validating syntax in real-time to prevent errors.
- **Group By & Aggregation:** Summarize your data by grouping it by one or more fields. Apply aggregate functions (COUNT, SUM, AVG, MIN, MAX) to your output fields to create powerful summary reports.
- **Flexible Report Design:** Select desired output fields from all available tables. Reorder fields to define the final report structure.

![image](https://github.com/user-attachments/assets/2dc15ff8-0935-4784-be35-5062ea0a1b87)

*Main Visual Designer with the new Condition Builder and Group By section.*

#### Simple Query Tab
- **Natural Language Queries:** For quick filtering, use a simple, intuitive syntax like `show name, city where state is 'CA' and year of order_date > 2022`.
- **Intellisense:** Get auto-completion suggestions for field names as you type.
- **Supports:** CONTAINS, STARTS WITH, ENDS WITH, numeric comparisons, and date functions (year of, month of, day of).

#### SQL View Tab
- View a SQL representation of the visually designed query.
- Toggle "Edit Mode" to write and execute custom SQL queries directly, enabling complex operations like UNION.

### Query Options & Results
- **Join Types:** Choose between Inner Join and Left Anti-Join.
- **Limit Results:** Optionally limit the number of returned rows.
- **Interactive Results Grid:**
  - Sort results by clicking any column header.
  - Navigate with "Go to Row" and "Next" buttons.
  - Export results to CSV.

### Save/Load Query
- Save the entire state of the designer (all tabs, tables, conditions, and output settings) to a .json file via Query > Save Query....
- Load a previously saved query configuration to instantly restore your work.

### User Experience
- **Built-in Help:** Press F1 to open a comprehensive help window with a clickable table of contents.
- **Keyboard Shortcuts:**
  - F1: Open Help
  - Ctrl+W: Resize result columns to fit content.
  - Ctrl+E: Export results to CSV.
- **Responsive UI:** Adjustable layout with paned windows.
- Error handling and informative messages.

## Requirements
- Python 3.6+
- Tkinter (usually included with standard Python installations)

## How to Run
1. Create a help.md file in the same directory as the script (content provided in a separate file).
2. Save the code as a Python file (e.g., xml_notepad.py).
3. Open a terminal or command prompt.
4. Navigate to the directory where you saved the files.
5. Run the script using:
```bash
python xml_notepad.py
```

## Usage Guide

### 1. Opening an XML File
- Go to File > Open XML.... The application will scan the file for tables.

### 2. Working with Detected Tables
- If tables are detected, they will be listed in the "Tables:" combobox. Select one to view its data in a grid.

### 3. Advanced Query Designer Usage
- Go to Utils > Query Designer....

#### Visual Designer Tab
1. **Select Tables:** Choose your Left (T1) and optionally a Right (T2) table.
2. **Define Conditions:**
   - If joining tables, select the Join (ON) radio button to define join keys.
   - Select the Filter (WHERE) radio button to define filter criteria.
   - Click Add to add the defined condition to the list.
   - Use the ( ), AND, OR, NOT buttons to build your logic. The list will validate your syntax in real-time.
3. **Group By (Optional):**
   - Move fields from "Available Fields" to "Group By Fields" to group your data.
4. **Design Report Output:**
   - If grouping, select an aggregate function (COUNT, SUM, etc.) from the dropdown.
   - Select fields from "Available Fields" and click Add Field > to move them to the "Selected Fields" list.
   - Use the < Remove, Up, and Down buttons to finalize your report columns.
5. **Run Query:** Click the Run Designer Query button to execute and see the results.

### Other Tabs & Features
- **Simple Query / SQL View:** Use these tabs for text-based querying as described in the Features section.
- **Save/Load:** Use the Query menu within the designer to save and load your work.
- **Help:** Press F1 at any time to open the detailed help document.

![image](https://github.com/user-attachments/assets/ba8b4fdd-8160-4683-8cd3-b43c05695684)


## Configuration Constants (in code)
- `CHUNK_SIZE = 1024 * 1024`: Size of chunks (1MB) for reading large XML files.
- `MIN_ROWS_FOR_TABLE = 3`: Minimum number of similar child elements required to be considered a potential table.
- `MIN_PERCENT_SIMILAR = 0.6`: Minimum percentage of children under a parent that must share the most common tag for table detection.

## Limitations & Known Issues
- **Performance:** While optimized, extremely large files may still cause UI lag. The query designer currently loads all data for selected tables into memory.
- **Manual Query Parsing:** The "Apply to Designer" feature in the SQL View works best with simple queries and may not correctly parse complex, nested SQL.
- **Data Modification:** This tool is focused on data querying and analysis. Direct editing of XML data is not supported.

## Future Enhancements (Ideas)
- **Performance:** Stream large tables from disk during query execution instead of loading them into memory.
- **Data Visualization:** Add a "Chart View" to create simple bar or pie charts from aggregated query results.
- **XPath Searching:** Implement a dedicated XPath search tool for advanced node selection.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
This project is open-source and released under the MIT License.

(c) Shuvro Basu, 2025.
