# Advanced Query Designer Help

This document explains how to use the different features of the Advanced Query Designer.

## Table of Contents
*   Visual Designer Tab
*   Simple Query Tab
*   SQL View Tab
*   Results Grid
*   Menu and Shortcuts

## Visual Designer Tab

The Visual Designer allows you to build queries graphically without writing SQL.

### 1. Select Tables
-   **Left Table (T1):** The primary table for your query. This is mandatory.
-   **Right Table (T2):** An optional second table to join with T1. Selecting a T2 enables join-based queries.

### 2. Define Conditions
This section lets you build the `WHERE` (for single tables) or `ON` (for joins) clause of your query. Conditions are built sequentially in the listbox.

-   **Add Condition:** After selecting fields and values in the input boxes above, click this to add a filter or join condition to the list.
-   **Logical Buttons:**
    -   `(`, `)`: Add opening and closing parentheses to create nested logical groups.
    -   `AND`, `OR`: Combine conditions.
    -   `NOT`: Negates the next condition or group.
-   **Workflow:**
    1.  Define a condition (e.g., `address_state = 'CH'`).
    2.  Click `Add Condition`.
    3.  Click a logical operator like `AND`. This clears the input fields.
    4.  Define the next condition and add it.
    5.  Example list: `(`, `Condition A`, `OR`, `Condition B`, `)`, `AND`, `Condition C`

### 3. Design Report Output
-   **Available Fields:** A list of all columns from the selected table(s).
-   **Shuttle Buttons (`>`, `>>`, `<`, `<<`):** Move fields between the "Available" and "Selected" lists.
-   **Selected Fields:** The fields that will appear in your results, in the specified order.
-   **Order Buttons (`Up`, `Down`):** Reorder the fields in the "Selected" list.

## Simple Query Tab

For users who prefer a simplified text-based language.

### Syntax
The basic syntax is `show [fields] where [conditions]`.

-   `show`: This keyword is mandatory.
-   `[fields]`: A comma-separated list of field names (e.g., `name, city, state`). Use `*` or leave this blank to select all fields.
-   `where`: This keyword is optional. If used, it must be followed by one or more conditions.

### Supported Operators
-   **String:** `CONTAINS`, `NOT CONTAINS`, `STARTS WITH`, `ENDS WITH`, `IS` (`=`), `IS NOT` (`!=`).
-   **Numeric:** `=`, `!=`, `>`, `<`, `>=`, `<=`.
-   **Date Functions:** `year of [field]`, `month of [field]`, `day of [field]`. Example: `where year of order_date > 2022`.

## SQL View Tab

Provides full control by allowing you to write and execute raw SQL queries.

-   **Run SQL:** Executes the query in the text box directly. This is independent of the Visual Designer. It supports more complex queries like `UNION`.
-   **Toggle Edit Mode:** Unlocks the text box for editing. The query shown is generated from the Visual Designer.
-   **Apply to Designer:** Parses the SQL in the text box and attempts to apply it to the Visual Designer. This works best for simple `SELECT` statements.

## Results Grid

Displays the output of your query.

-   **Sorting:** Click any column header to sort the results by that column. Click again to reverse the sort order.
-   **Navigation:**
    -   `Go to Row`: Enter a row number and press Enter to jump to that row.
    -   `Next`: Moves the selection to the next row in the grid.

## Menu and Shortcuts

-   **Query > Load/Save Query:** Save or load the complete state of the designer (all tabs) to a `.json` file.
-   **Query > Export Results:** Save the current results grid as a `.csv` file.
-   **F1:** Opens this help window.
-   **Ctrl+W:** Resizes result columns to fit their content.
-   **Ctrl+E:** Exports results to CSV.