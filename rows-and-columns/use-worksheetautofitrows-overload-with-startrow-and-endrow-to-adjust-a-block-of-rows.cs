// Title: How to auto‑fit rows 1‑5 in an Aspose.Cells .NET worksheet using Worksheet.AutoFitRows(startRow, endRow)
// AI Prompts: Write C# code that applies text wrapping to a range and then calls Worksheet.AutoFitRows with specific start and end row indices. | Show a complete example that creates a workbook, fills rows with long text, enables wrapping, auto‑fits rows 0 to 4, and saves the file. | Demonstrate using the AutoFitRows overload to adjust the height of a selected block of rows in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# auto fit rows 1 to 5 example | Worksheet.AutoFitRows startRow endRow parameters explanation | How to auto adjust row height for a specific range in Aspose.Cells .NET | Applying text wrap style before using AutoFitRows in Aspose.Cells
// Tags: Worksheet.AutoFitRows range height adjustment | apply text wrap before row height auto‑adjust | C# Aspose.Cells row height tuning | selective row height scaling with start and end indices | save workbook after row height adjustment

using System;
using Aspose.Cells;

// Creates a workbook, populates rows with long text, applies a text‑wrap style to A1:C5, uses Worksheet.AutoFitRows(0, 4) to auto‑fit rows 1‑5, and saves the file as AutoFitRowsExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Access the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data in rows 1 to 5 (zero‑based indices 0‑4).
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                // Insert a long text to demonstrate row height adjustment.
                sheet.Cells[row, col].PutValue($"Row {row + 1}, Column {col + 1}: This is a sample text that may require wrapping.");
            }
        }

        // Enable text wrapping so that AutoFitRows can increase row height if needed.
        Style style = workbook.CreateStyle();
        style.IsTextWrapped = true;
        StyleFlag flag = new StyleFlag() { WrapText = true };
        sheet.Cells.CreateRange("A1:C5").ApplyStyle(style, flag);

        // Auto‑fit rows from the first row (index 0) to the fifth row (index 4).
        sheet.AutoFitRows(0, 4);

        // Save the workbook to a file.
        workbook.Save("AutoFitRowsExample.xlsx");
    }
}
