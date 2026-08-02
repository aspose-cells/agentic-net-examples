// Title: Enable Excel 2003 Compatibility for a PivotTable in Aspose.Cells (C#) – Truncate Text >255 Characters
// Description: This C# example creates a workbook, adds rows with a description longer than 255 characters, builds a pivot table, sets PivotTable.IsExcel2003Compatible to true, then calls RefreshData and CalculateData. The setting forces Excel 2003‑style truncation of strings over 255 characters before the pivot is refreshed, and the workbook is saved as an .xlsx file.
// Keywords: Aspose.Cells | PivotTable | IsExcel2003Compatible | C# | .NET | Excel 2003 compatibility | truncate long text | 255 characters limit | RefreshData | CalculateData | sample code | GitHub | pivot table truncation
// Common Searches: Aspose.Cells PivotTable truncate text 255 characters | Set IsExcel2003Compatible before RefreshData | C# example for Excel 2003 compatible pivot table | How to limit pivot table strings to 255 chars in Aspose.Cells | PivotTable.IsExcel2003Compatible usage
// Developer Intent: Configure a pivot table to operate in Excel 2003 compatibility mode so that any string longer than 255 characters is automatically truncated before the pivot data is refreshed.
// Use Cases: Processing legacy Excel files that require the 255‑character string limit. | Generating reports with long description fields while ensuring older Excel versions can open the pivot without errors. | Demonstrating the required order of setting IsExcel2003Compatible and calling RefreshData/CalculateData.
// AI Prompts: Generate C# code using Aspose.Cells that creates a pivot table, sets IsExcel2003Compatible to true before RefreshData, and saves the workbook. | Explain why PivotTable.IsExcel2003Compatible must be assigned before RefreshData and show the effect on long text values. | Provide a step‑by‑step tutorial for enabling Excel 2003 compatibility on a pivot table and verifying the 255‑character truncation.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // This C# example creates a workbook, adds rows with a description longer than 255 characters, builds a pivot table, sets PivotTable.IsExcel2003Compatible to true, then calls RefreshData and CalculateData. The setting forces Excel 2003‑style truncation of strings over 255 characters before the pivot is refreshed, and the workbook is saved as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and add sample data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].Value = "Product";
            dataSheet.Cells["B1"].Value = "Description";

            // Data rows (including a long description >255 chars)
            dataSheet.Cells["A2"].Value = "Item1";
            dataSheet.Cells["B2"].Value = new string('X', 300); // long text

            dataSheet.Cells["A3"].Value = "Item2";
            dataSheet.Cells["B3"].Value = "Short description";

            // Add a worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Add a pivot table based on the data range A1:B3, place it at A4
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B3", "A4", "MyPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure fields: Product as row, Description as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);    // Description column

            // Enforce Excel 2003 compatibility (truncate strings >255 chars)
            pivotTable.IsExcel2003Compatible = true;

            // Refresh data and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_Excel2003Compatible.xlsx");
        }
    }
}
