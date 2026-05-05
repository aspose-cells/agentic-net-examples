using System;
using System.Data;
using Aspose.Cells;

class SmartMarkerGroupingDemo
{
    static void Main()
    {
        // Load the template workbook that contains smart markers.
        // The template should have a named range "_CellsSmartMarkers" covering the data row,
        // e.g., cells A2:C2 with markers &=$Name, &=$Category, &=$Amount.
        Workbook workbook = new Workbook("Template.xlsx");

        // Prepare a DataTable that will be bound to the smart markers.
        DataTable salesTable = new DataTable("Sales");
        salesTable.Columns.Add("Name", typeof(string));
        salesTable.Columns.Add("Category", typeof(string));
        salesTable.Columns.Add("Amount", typeof(double));

        salesTable.Rows.Add("Alice",   "East",  1200);
        salesTable.Rows.Add("Bob",     "West",  1500);
        salesTable.Rows.Add("Charlie", "East",   800);
        salesTable.Rows.Add("David",   "North",  950);
        salesTable.Rows.Add("Eve",     "West",  1100);

        // Set up the WorkbookDesigner and bind the data source.
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;
        designer.SetDataSource("Sales", salesTable);

        // Process all smart markers in the workbook.
        // The boolean parameter indicates whether to preserve unrecognized markers.
        designer.Process(true);

        // After processing, group rows by the "Category" column.
        // For demonstration, we sort the DataTable so that rows of the same category are contiguous.
        salesTable.DefaultView.Sort = "Category ASC";
        // Re‑process to reflect the sorted order (optional, depending on template design).
        // Here we simply re‑process the same data.
        designer.Process(true);

        // The processed data starts at row 2 (zero‑based index 1).
        // Group the contiguous rows for each category.
        // Adjust the indices if the actual row count differs.
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // East category rows (2‑3)
        cells.GroupRows(1, 2, true);
        // West category rows (4‑5)
        cells.GroupRows(3, 4, true);
        // North category is a single row; grouping is not required.

        // Save the final workbook with grouped rows.
        workbook.Save("GroupedReport.xlsx");
    }
}