// Title: Enable Excel 2003 Compatibility for PivotTables in Aspose.Cells (C#)
// Description: This example shows how to create a workbook, add sample data, build a PivotTable, set the IsExcel2003Compatible property to true, and then refresh and calculate the PivotTable so that any text longer than 255 characters is automatically truncated for Excel 2003 compatibility. The workbook is saved as an .xlsx file.
// Keywords: Aspose.Cells | C# PivotTable | Excel 2003 compatibility | IsExcel2003Compatible | truncate long strings | RefreshData | CalculateData | pivot table compatibility mode | .NET Excel library
// Common Searches: Aspose.Cells enable Excel 2003 compatibility for pivot table | IsExcel2003Compatible property C# example | truncate strings over 255 characters in Aspose.Cells pivot | refresh pivot table after setting compatibility mode | how to make pivot table Excel 2003 compatible using Aspose
// Developer Intent: Set Excel 2003 compatibility on a PivotTable before refreshing it with Aspose.Cells.
// Use Cases: Generate a report that must open in Excel 2003 without errors caused by long text fields in a PivotTable. | Create automated Excel files where PivotTable data is refreshed programmatically while ensuring legacy compatibility. | Process product catalogs with descriptions exceeding 255 characters and automatically truncate them in the PivotTable output.
// AI Prompts: Write C# code that creates a workbook, adds data, builds a PivotTable, enables Excel 2003 compatibility, refreshes the pivot, and saves the file using Aspose.Cells. | Explain the effect of the IsExcel2003Compatible property on string length handling in Aspose.Cells PivotTables. | Provide step‑by‑step instructions to configure a PivotTable for Excel 2003 compatibility, refresh it, and export the workbook in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExcel2003Compatibility
{
    // This example shows how to create a workbook, add sample data, build a PivotTable, set the IsExcel2003Compatible property to true, and then refresh and calculate the PivotTable so that any text longer than 255 characters is automatically truncated for Excel 2003 compatibility. The workbook is saved as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and add sample data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";
            dataSheet.Cells["A1"].Value = "Product";
            dataSheet.Cells["B1"].Value = "Description";
            dataSheet.Cells["C1"].Value = "Quantity";

            dataSheet.Cells["A2"].Value = "Item1";
            dataSheet.Cells["B2"].Value = "Short description";
            dataSheet.Cells["C2"].Value = 10;

            dataSheet.Cells["A3"].Value = "Item2";
            dataSheet.Cells["B3"].Value = "Very long description that would exceed the 255‑character limit in Excel 2003 when used in a pivot table. This text is intentionally long to demonstrate the compatibility setting.";
            dataSheet.Cells["C3"].Value = 20;

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create a pivot table based on the data range
            int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:C3", "A5", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Quantity as data field

            // Enable Excel 2003 compatibility before refreshing
            // This ensures that any string longer than 255 characters will be truncated during refresh.
            pivotTable.IsExcel2003Compatible = true;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_Excel2003Compatibility.xlsx");
        }
    }
}
