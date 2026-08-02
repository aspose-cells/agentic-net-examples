// Title: Freeze Header Row in Excel with Aspose.Cells for .NET (C#) and Save as New XLSX
// Description: Loads an existing workbook (input.xlsx), accesses the first worksheet, applies FreezePanes at cell A2 to lock the top row, and saves the result to a new file (output.xlsx) using Aspose.Cells for .NET.
// Keywords: Aspose.Cells FreezePanes C# | freeze top row Excel .NET | lock header row Aspose.Cells | save workbook as new XLSX | Aspose.Cells worksheet freeze example
// Common Searches: how to freeze the first row in Excel with Aspose.Cells C# | Aspose.Cells FreezePanes example for header row | C# code to freeze top row and save new file | Aspose.Cells freeze panes and export workbook
// Developer Intent: The developer needs to programmatically freeze the first (header) row of a worksheet and write the modified workbook to a separate XLSX file.
// Use Cases: Create reports where column titles stay visible while scrolling through large data sets. | Distribute Excel templates that keep header information fixed for collaborators. | Export data extracts with frozen headers to improve usability for downstream analysts.
// AI Prompts: Generate C# code to freeze multiple rows and columns with Aspose.Cells FreezePanes. | Show how to unfreeze panes and reset the view in an Aspose.Cells workbook. | Explain how to determine the freeze range dynamically based on variable header rows using Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an existing workbook (input.xlsx), accesses the first worksheet, applies FreezePanes at cell A2 to lock the top row, and saves the result to a new file (output.xlsx) using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Freeze the first row (header row)
        // FreezePanes(cellName, freezedRows, freezedColumns)
        // "A2" means the freeze line is just below the first row.
        worksheet.FreezePanes("A2", 1, 0);

        // Save the modified workbook to a new XLSX file
        workbook.Save("output.xlsx");
    }
}
