// Title: Set PivotTable.IsExcel2003Compatible = true to truncate strings >255 characters in Aspose.Cells (C#)
// Description: Creates a workbook with a column containing a 300‑character string, adds a pivot table on a separate sheet, enables Excel 2003 compatibility by setting PivotTable.IsExcel2003Compatible to true, refreshes and calculates the pivot so values longer than 255 characters are automatically truncated, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | PivotTable | IsExcel2003Compatible | truncate long strings | 255 character limit | C# | .NET | Excel 2003 compatibility | RefreshData | CalculateData | pivot cache
// Common Searches: How to truncate values longer than 255 characters in an Aspose.Cells pivot table | Set IsExcel2003Compatible before RefreshData in C# | Enable Excel 2003 compatibility for Aspose.Cells pivot tables | Why does my Aspose pivot table show full text instead of truncated | Aspose.Cells pivot table string length limit
// Developer Intent: Enable Excel 2003 compatibility on an Aspose.Cells PivotTable so that any field value exceeding 255 characters is truncated before the pivot is refreshed.
// Use Cases: Generate a report where description fields may exceed 255 characters but the pivot must follow Excel 2003 limits. | Create a pivot table on a separate worksheet, enforce compatibility, then refresh and calculate to apply truncation automatically. | Save an Excel file that mimics legacy Excel 2003 behavior for downstream systems that require the 255‑character restriction.
// AI Prompts: Write C# code using Aspose.Cells that builds a pivot table, sets IsExcel2003Compatible to true, refreshes the pivot, and saves the workbook. | Explain the effect of PivotTable.IsExcel2003Compatible on string length handling in Aspose.Cells and show how to verify truncation. | Provide a step‑by‑step tutorial for truncating long text fields in a pivot table by enabling Excel 2003 compatibility before refreshing.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook with a column containing a 300‑character string, adds a pivot table on a separate sheet, enables Excel 2003 compatibility by setting PivotTable.IsExcel2003Compatible to true, refreshes and calculates the pivot so values longer than 255 characters are automatically truncated, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet for source data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate the worksheet with sample data, including a long string (>255 chars)
        dataSheet.Cells["A1"].Value = "Product";
        dataSheet.Cells["B1"].Value = "Description";

        dataSheet.Cells["A2"].Value = "Item1";
        dataSheet.Cells["B2"].Value = new string('X', 300); // 300‑character string

        // Add a new worksheet that will contain the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create a pivot table using the source range A1:B2 and place it starting at cell A4
        int pivotIndex = pivotSheet.PivotTables.Add("A1:B2", "A4", "MyPivotTable");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Add fields to the pivot table (Product as row, Description as data)
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Product
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Description

        // Enforce Excel 2003 compatibility so that strings longer than 255 characters are truncated
        pivotTable.IsExcel2003Compatible = true;

        // Refresh the pivot cache and calculate the pivot data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook to a file
        workbook.Save("PivotExcel2003Compatible.xlsx");
    }
}
