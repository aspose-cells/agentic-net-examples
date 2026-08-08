// Title: Copy Rows Values‑Only Using PasteOptions in Aspose.Cells for .NET (C#)
// Description: Shows how to copy a source range of rows to another location in an Aspose.Cells workbook using PasteOptions set to transfer only calculated values, skip blanks, and ignore formulas and formatting. The C# example saves the file as CopyRowsValuesOnly.xlsx.
// Keywords: Aspose.Cells | PasteOptions | Values only | C# | .NET | copy rows | ignore formulas | skip blanks | range copy | static data export | Excel automation
// Common Searches: Aspose.Cells copy only values C# | PasteOptions values only example | Copy rows without formulas Aspose.Cells | Skip blanks when copying range Aspose | C# copy range values only Excel
// Developer Intent: Transfer rows from a source range to a destination while preserving only the resulting cell values and discarding formulas, formatting, and external links.
// Use Cases: Generate a report sheet with static results for distribution | Export calculated data to another workbook without exposing formulas | Create an archival snapshot of a calculation block with plain values | Prepare data for third‑party systems that require raw values only
// AI Prompts: Show how to keep number formatting while copying only values with PasteOptions. | Provide code to copy values only then remove formulas from the source range. | Explain using PasteOptions to copy only visible cells after applying a filter. | Demonstrate copying rows to a different workbook while preserving values only.

using System;
using Aspose.Cells;

// Shows how to copy a source range of rows to another location in an Aspose.Cells workbook using PasteOptions set to transfer only calculated values, skip blanks, and ignore formulas and formatting. The C# example saves the file as CopyRowsValuesOnly.xlsx.
class CopyRowsValuesOnly
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source rows with values and formulas
            // Row 0
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].Formula = "=A1*2";
            // Row 1
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B2"].Formula = "=A2*2";
            // Row 2
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["B3"].Formula = "=A3*2";

            // Define the source range (first three rows, columns A:B)
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange(0, 0, 3, 2);

            // Define the destination range (starting at row 5, columns A:B)
            Aspose.Cells.Range destRange = sheet.Cells.CreateRange(5, 0, 3, 2);

            // Configure PasteOptions to copy only values (ignore formulas and formatting)
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.Values,   // copy only cell values
                SkipBlanks = true,              // skip blank cells
                OnlyVisibleCells = false,       // copy all cells regardless of visibility
                Transpose = false,              // no transposition
                OperationType = PasteOperationType.None,
                IgnoreLinksToOriginalFile = true
            };

            // Perform the copy with the specified paste options
            destRange.Copy(sourceRange, pasteOptions);

            // Save the workbook to verify the result
            workbook.Save("CopyRowsValuesOnly.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
