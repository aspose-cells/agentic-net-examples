// Title: Hide Columns 5‑8, Unhide 6‑7 with Width, and Export to PDF using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, hides columns 5‑8, unhides columns 6‑7 while setting their width to 15 points, and saves the worksheet as a PDF file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells hide columns C# | Aspose.Cells unhide columns with width | export worksheet to PDF Aspose.Cells | set column width Aspose.Cells | column visibility PDF export
// Common Searches: Aspose.Cells hide columns 5 to 8 C# | how to unhide columns 6 and 7 and set width Aspose.Cells | export Excel to PDF after changing column visibility .NET | C# code to hide and unhide columns with Aspose.Cells | set column width when unhiding columns Aspose.Cells
// Developer Intent: Hide columns 5‑8, then unhide columns 6‑7 with a specific width, and generate a PDF.
// Use Cases: Create a printable PDF where sensitive columns are hidden but selected columns remain visible with a defined width. | Automate report generation that requires temporary column hiding before exporting to PDF. | Adjust column layout programmatically for consistent PDF formatting in financial or inventory reports.
// AI Prompts: Generate C# code using Aspose.Cells to hide columns 5‑8, unhide columns 6‑7 with a width of 15 points, and save the workbook as a PDF. | Explain how HideColumns and UnhideColumns work in Aspose.Cells and how to specify column width when unhiding. | Show an example of exporting a worksheet to PDF after modifying column visibility and widths with Aspose.Cells for .NET.

using Aspose.Cells;

// Loads an Excel workbook, hides columns 5‑8, unhides columns 6‑7 while setting their width to 15 points, and saves the worksheet as a PDF file with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide columns 5 through 8 (human numbers) -> zero‑based indices 4 to 7
        cells.HideColumns(4, 4); // start at index 4, hide 4 columns

        // Unhide columns 6 and 7 (human numbers) -> indices 5 and 6, total 2 columns, set width to 15.0
        cells.UnhideColumns(5, 2, 15.0);

        // Export the worksheet to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
