// Title: Aspose.Cells .NET – Enable PivotTable Auto‑Refresh on Workbook Open
// Description: C# example that creates a workbook, adds sample sales data, builds a PivotTable, sets the RefreshDataOnOpeningFile property to true, and saves the file so the pivot updates automatically each time the workbook is opened.
// Keywords: Aspose.Cells | PivotTable auto refresh | RefreshDataOnOpeningFile | C# | .NET | Excel workbook open | pivot refresh on load | Aspose.Cells PivotTable property
// Common Searches: Aspose.Cells set PivotTable refresh on open | RefreshDataOnOpeningFile C# example | auto refresh pivot table Aspose.Cells .NET | make PivotTable update when workbook is opened | Aspose.Cells PivotTable auto‑refresh property
// Developer Intent: Configure a PivotTable to refresh automatically when the Excel file is opened using Aspose.Cells for .NET.
// Use Cases: Distribute a sales report that always reflects the latest data without manual refresh. | Create a financial dashboard where pivot calculations are refreshed on every file load. | Automate periodic data exports that generate Excel workbooks with self‑updating PivotTables.
// AI Prompts: Generate C# code that builds a workbook, adds data, creates a PivotTable, enables RefreshDataOnOpeningFile, and saves the file with Aspose.Cells. | Explain the purpose of the RefreshDataOnOpeningFile property for PivotTables in Aspose.Cells and best practices for its use. | Show how to apply auto‑refresh to multiple PivotTables in a single workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshOnOpenDemo
{
    // C# example that creates a workbook, adds sample sales data, builds a PivotTable, sets the RefreshDataOnOpeningFile property to true, and saves the file so the pivot updates automatically each time the workbook is opened.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1000);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(2000);
            worksheet.Cells["A4"].PutValue("Orange");
            worksheet.Cells["B4"].PutValue(3000);

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure the pivot table (row field and data field)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales column

            // Enable auto‑refresh when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = true;

            // Save the workbook
            workbook.Save("PivotTable_AutoRefreshOnOpen.xlsx");
        }
    }
}
