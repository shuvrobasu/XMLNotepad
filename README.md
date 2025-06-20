# XML/CSV-Notepad - Version 4.1.0

XML/CSV-Notepad is a Python-based desktop application built with Tkinter for viewing and navigating XML and CSV files. It excels at identifying, querying, and working with tabular data structures. Users can open large files, have tables automatically detected, and perform complex data operations using an intuitive, multi-tabbed interface and a powerful query designer.

![image](https://github.com/user-attachments/assets/7350a18c-6da4-4a2b-9c09-434fd78da8e3)


## Key Features

### XML/CSV Viewing & Navigation
- Open and parse large XML or CSV files efficiently.
- Display XML structure in a hierarchical tree view.
- View details of selected XML nodes (tag, text, attributes).

### Automatic Table Detection & Tabbed Interface
- Intelligently identifies and extracts tabular data from XML structures.
- **Multi-Tabbed Viewing:** Open multiple tables in a dynamic tabbed interface. Each table appears in its own tab, allowing for easy comparison and multitasking.
- **Click-to-Open:** Click a table node in the XML tree to instantly open its data in a new tab.
- **Tab Management:** Easily close tabs using the x button on each tab, or reorder them using the Edit > Reorder Tabs... dialog.

### Data Integrity & Manipulation
- **Find and Replace:** A powerful dialog (Edit > Find...) to find text within tables and perform single or batch replacements. This operation is fully supported by the undo/redo system.
- **Batch Operations:** A dedicated UI (Edit > Batch Operations...) to perform bulk UPDATE or DELETE actions on rows that match specific filter criteria. Includes a preview of affected rows and is undoable.
- **Transactional Data Check:** A utility (Utils > Transactional Data Check...) to ensure data integrity by finding orphaned records (e.g., invoices without a customer) or unused master records (e.g., customers with no invoices).

![image](https://github.com/user-attachments/assets/039edd69-e48a-4777-a4a5-4df18567b3bc)

![image](https://github.com/user-attachments/assets/e7fd4d61-5fa7-4640-b48c-ce6f700a5e92)  ![image](https://github.com/user-attachments/assets/55852eb3-8e7b-4ace-a904-8dc2387c07d4)

### Data Analysis & Querying
- **Column Statistics:** Right-click any table header to get instant statistics, including count, unique values, sum, and average for that column's data.
- **Advanced Query Designer:** A comprehensive tool for building complex queries.
  - **Visual Designer:** Build queries visually with support for joins, complex WHERE clauses (using parentheses and AND/OR/NOT), GROUP BY, and aggregate functions (COUNT, SUM, AVG, etc.).
  - **Simple Query:** Use a natural language syntax like `show name, city where state is 'CA'` for quick filtering.
  - **SQL View:** Write or view raw SQL queries for maximum flexibility.
  - **Save/Load Query:** Save and load entire query designer sessions, including all conditions and settings, to a JSON file.

![image](https://github.com/user-attachments/assets/60819234-d560-4582-8373-4ed4cc84e46e)


![alt text](https://github.com/user-attachments/assets/2dc15ff8-0935-4784-be35-5062ea0a1b87)

*The Advanced Query Designer allows for powerful, visually constructed data analysis.*

## User Experience
- **Built-in Help:** Press F1 to open a comprehensive help window.
- **Keyboard Shortcuts:** Use Ctrl+O to open, Ctrl+F for find/replace, and Ctrl+S to save the current table as a CSV.
- **Responsive UI:** Adjustable layout with paned windows.
- Error handling and informative messages, including a memory warning when many tabs are open.

## Requirements
- Python 3.6+
- Tkinter (usually included with standard Python installations)
- lxml (`pip install lxml`)

## How to Run
1. Create a help.md file in the same directory as the script.
2. Save the code as a Python file (e.g., xml_notepad.py).
3. Open a terminal or command prompt.
4. Navigate to the directory where you saved the files.
5. Run the script using:
```bash
python xml_notepad.py
```

## Usage Guide

### 1. Opening a File
- Go to File > Open XML... or File > Open CSV.... The application will scan the file for tables.

### 2. Working with Tables
- Detected tables are listed in the "Tables:" combobox.
- To open a table: Select it from the combobox or, for XML files, click on its parent node in the tree view on the left. The table will open in a new tab.
- To close a tab: Click the x button on the tab.
- To reorder tabs: Go to Edit > Reorder Tabs... and use the "Move Up" / "Move Down" buttons.

### 3. Analyzing & Manipulating Data
- **Column Statistics:** In any table tab, right-click a column header and select "Show Column Statistics" for an instant analysis.
- **Find and Replace:** Go to Edit > Find... (Ctrl+F). Enter your search and replace terms, then use "Find All", "Replace", or "Replace All".
- **Batch Operations:** Go to Edit > Batch Operations.... Define a filter to select rows, then choose to either update a specific column or delete all matching rows. Always preview your changes before executing.

### 4. Transactional Data Check
- Go to Utils > Transactional Data Check....
- Select your primary table (e.g., Students) and its key (student_id).
- Select your transactional table (e.g., Marks) and its foreign key (student_id).
- Choose whether to find orphaned transactions or unused primary records, then run the analysis.

### 5. Advanced Query Designer
- Go to Utils > Query Designer....
- Use the Visual Designer to build queries by selecting tables, defining conditions, grouping data, and designing the final report output.
- Click Run Designer Query to see the results.

![alt text](https://github.com/user-attachments/assets/ba8b4fdd-8160-4683-8cd3-b43c05695684)

## Configuration Constants (in code)
- `CHUNK_SIZE = 1024 * 1024`: Size of chunks (1MB) for reading large files.
- `MIN_ROWS_FOR_TABLE = 3`: Minimum number of similar child elements required to be considered a potential table.
- `MIN_PERCENT_SIMILAR = 0.6`: Minimum percentage of children under a parent that must share the most common tag for table detection.

## Limitations & Known Issues
- **Performance:** While optimized, extremely large files may still cause UI lag. The query designer and batch operations currently load data for selected tables into memory.
- **Manual Query Parsing:** The "Apply to Designer" feature in the SQL View works best with simple queries and may not correctly parse complex, nested SQL.

## Future Enhancements (Ideas)
- **Data Visualization:** Add a "Chart View" to create simple bar or pie charts from aggregated query results.
- **XPath Searching:** Implement a dedicated XPath search tool for advanced node selection in XML files.
- **Performance:** Stream large tables from disk during query execution instead of loading them fully into memory.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
This project is open-source and released under the MIT License.

(c) Shuvro Basu, 2025.
