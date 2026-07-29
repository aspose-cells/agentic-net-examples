// Title: Hide Columns 5‑8, Unhide 6‑7 with Width, Export to PDF using Aspose.Cells for .NET (C#)
// Description: Load a workbook, hide columns 5‑8, then unhide columns 6‑7 setting their width to 15 points, and save the worksheet as a PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells hide columns C# | Aspose.Cells unhide columns width | Aspose.Cells export to PDF | C# column visibility Aspose.Cells | set column width Aspose.Cells | HideColumns Aspose.Cells | UnhideColumns Aspose.Cells | Aspose.Cells PDF conversion
// Common Searches: Aspose.Cells hide columns 5 to 8 C# | Unhide specific columns with width Aspose.Cells | Export worksheet to PDF after hiding columns | How to set column width while unhiding in Aspose.Cells | C# hide and unhide columns before PDF export
// Developer Intent: Hide columns 5‑8, unhide columns 6‑7 with a 15‑point width, and generate a PDF file.
// Use Cases: Create a printable report that omits internal calculation columns while showing key data columns at a defined width. | Prepare a PDF invoice where confidential columns are hidden but essential columns are displayed with proper sizing. | Generate a clean PDF version of a spreadsheet for distribution, adjusting column visibility and width before export.
// AI Prompts: Write C# code with Aspose.Cells to hide columns 5‑8, then unhide columns 6‑7 setting their width to 15 points and save as PDF. | Show an Aspose.Cells example that demonstrates hiding a range of columns, adjusting width on unhide, and exporting the sheet to PDF. | Explain the zero‑based column indexing in Aspose.Cells and how to use HideColumns and UnhideColumns with a width parameter.

using System;
using Aspose.Cells;

// Load a workbook, hide columns 5‑8, then unhide columns 6‑7 setting their width to 15 points, and save the worksheet as a PDF with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide columns 5 through 8 (zero‑based indices 4,5,6,7)
        cells.HideColumns(4, 4);

        // Unhide columns 6 and 7 (indices 5 and 6) and set their width to 15.0
        cells.UnhideColumns(5, 2, 15.0);

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
