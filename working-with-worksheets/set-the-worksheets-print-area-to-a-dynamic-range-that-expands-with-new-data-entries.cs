// Title: Define a dynamic print area that expands with data in an Aspose.Cells worksheet (C#)
// Common Searches: C# Aspose.Cells set print area to range of used cells | How to automatically expand worksheet print area when adding rows in Aspose.Cells | Aspose.Cells PageSetup.PrintArea based on MaxDataRow and MaxDataColumn | Set print area to last populated cell in an Excel workbook using Aspose.Cells | Dynamic print range for Excel files with Aspose.Cells .NET
// Tags: Aspose.Cells dynamic worksheet print area | PageSetup.PrintArea using MaxDataRow | auto expand Excel print range C# | calculate used range for print area Aspose.Cells | programmatic worksheet print area definition

using Aspose.Cells;
using System;

// The example creates a workbook, populates sample data, determines the last occupied row and column with MaxDataRow and MaxDataColumn, builds an A1‑style address from A1 to that cell, assigns the range to Worksheet.PageSetup.PrintArea, and saves the file, resulting in a print area that automatically grows as new data is added.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data (replace with your actual data population logic)
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue(100);
        sheet.Cells["B2"].PutValue(200);
        sheet.Cells["A3"].PutValue(300);
        sheet.Cells["B3"].PutValue(400);

        // Determine the last row and column that contain data
        int lastRow = sheet.Cells.MaxDataRow;      // zero‑based index
        int lastCol = sheet.Cells.MaxDataColumn;   // zero‑based index

        // Build the A1‑style address for the dynamic print area
        string startAddress = sheet.Cells[0, 0].Name;                 // "A1"
        string endAddress   = sheet.Cells[lastRow, lastCol].Name;    // e.g., "B3"

        // Set the worksheet's print area to the calculated range
        sheet.PageSetup.PrintArea = $"{startAddress}:{endAddress}";

        // Save the workbook
        workbook.Save("DynamicPrintArea.xlsx");
    }
}
