// Title: Select a Worksheet by Index or Name with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a Workbook, add sheets, access a sheet using its zero‑based position or its assigned name, rename a sheet, write values to cells, and save the result as an XLSX file using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | select worksheet | worksheet index | worksheet name | Workbook.Worksheets | rename sheet | write cell values | save XLSX
// Common Searches: Aspose.Cells get worksheet by index C# | Aspose.Cells retrieve worksheet by name .NET | how to rename a worksheet using Aspose.Cells | write data to a specific sheet Aspose.Cells C# | save workbook as XLSX with Aspose.Cells
// Developer Intent: Retrieve a specific worksheet from a workbook via the Worksheets collection.
// Use Cases: Populate header rows on the first sheet (index 0) before importing data. | Insert a generation timestamp into a "Summary" sheet identified by its name. | Rename a newly added sheet to reflect its purpose before exporting the file.
// AI Prompts: Show C# code that accesses a worksheet by index and writes header values with Aspose.Cells. | Provide an example of selecting a worksheet by its name, updating cells, and saving the workbook as XLSX using Aspose.Cells for .NET. | Explain how to loop through all worksheets in a workbook and perform custom actions on each with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a Workbook, add sheets, access a sheet using its zero‑based position or its assigned name, rename a sheet, write values to cells, and save the result as an XLSX file using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook (default contains one worksheet)
        Workbook workbook = new Workbook();

        // Add additional worksheets with specific names
        workbook.Worksheets.Add("SalesData");
        workbook.Worksheets.Add("Summary");

        // ---- Select a worksheet by index ----
        // Index is zero‑based; this gets the first worksheet (original one)
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.Name = "Data"; // rename for clarity

        // ---- Select a worksheet by name ----
        // Retrieves the worksheet whose Name property matches "Summary"
        Worksheet summarySheet = workbook.Worksheets["Summary"];

        // Demonstrate that the correct sheets are accessed by writing data
        firstSheet.Cells["A1"].PutValue("Item");
        firstSheet.Cells["B1"].PutValue("Quantity");

        summarySheet.Cells["A1"].PutValue("Report generated on:");
        summarySheet.Cells["B1"].PutValue(DateTime.Now);

        // Save the workbook to disk
        workbook.Save("SelectedSheetsDemo.xlsx", SaveFormat.Xlsx);
    }
}
