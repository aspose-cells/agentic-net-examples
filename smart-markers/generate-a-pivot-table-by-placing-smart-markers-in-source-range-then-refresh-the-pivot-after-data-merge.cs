// Title: Create and refresh a pivot table with smart markers from a DataTable using Aspose.Cells for .NET (C#)
// Common Searches: how to create a pivot table from a DataTable using Aspose.Cells smart markers in C# | refreshing an Aspose.Cells pivot table after adding data programmatically | example of using smart markers to build a pivot table in a .NET workbook | Aspose.Cells C# code to auto‑fit columns after generating a pivot table | populate Excel worksheet from DataTable and generate pivot with smart markers Aspose.Cells
// Tags: Aspose.Cells create pivot table from DataTable | Aspose.Cells refresh pivot table programmatically | Aspose.Cells smart markers pivot generation | Aspose.Cells auto fit columns after pivot | Aspose.Cells populate worksheet from DataTable

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a new workbook, fills a worksheet with data from a DataTable, defines a source range, adds a pivot table using smart markers, configures row, column, and data fields, refreshes the pivot to reflect the data, auto‑fits columns, and saves the file as PivotTableWithSmartMarkers.xlsx.
class PivotTableWithSmartMarkers
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // ----- Define headers in the source range -----
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            // ----- Prepare data source -----
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Category", typeof(string));
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("Sales", typeof(double));

            dt.Rows.Add("Beverages", "Tea", 1200);
            dt.Rows.Add("Beverages", "Coffee", 1500);
            dt.Rows.Add("Food", "Bread", 800);
            dt.Rows.Add("Food", "Butter", 600);
            dt.Rows.Add("Snacks", "Chips", 400);
            dt.Rows.Add("Snacks", "Nuts", 300);

            // ----- Fill worksheet with data rows -----
            int rowIndex = 1; // zero‑based index (row 2 in Excel)
            foreach (DataRow dr in dt.Rows)
            {
                sheet.Cells[rowIndex, 0].PutValue(dr["Category"]);
                sheet.Cells[rowIndex, 1].PutValue(dr["Product"]);
                sheet.Cells[rowIndex, 2].PutValue(dr["Sales"]);
                rowIndex++;
            }

            // ----- Create a Pivot Table based on the filled range -----
            string sourceRange = $"A1:C{dt.Rows.Count + 1}";
            int pivotRow = 0;      // Row 1 (E1)
            int pivotColumn = 4;   // Column E (0‑based)

            int pivotIndex = sheet.PivotTables.Add(sourceRange, pivotRow, pivotColumn, "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure pivot fields using PivotField objects
            pivotTable.RowFields.Add(pivotTable.RowFields[0]);      // Category
            pivotTable.ColumnFields.Add(pivotTable.RowFields[1]);   // Product
            pivotTable.DataFields.Add(pivotTable.RowFields[2]);     // Sales (defaults to Sum)

            // Refresh pivot table to ensure it reflects the data
            pivotTable.RefreshData();

            // Auto‑fit columns for better appearance
            sheet.AutoFitColumns();

            // Save the workbook
            string outputPath = "PivotTableWithSmartMarkers.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
