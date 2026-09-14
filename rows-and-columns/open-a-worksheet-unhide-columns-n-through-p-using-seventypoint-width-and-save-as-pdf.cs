// Title: How to unhide columns N‑P, set a 70‑point width, and export a worksheet to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Unhide columns N through P, set each column width to 70 points, and save the workbook as a PDF with Aspose.Cells in C#. | Load an existing .xlsx, reveal columns 14‑16, adjust their width to 70 points, then export the sheet to PDF using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# unhide columns N to P and set column width in points | Set column width to 70 points before exporting Excel to PDF with Aspose.Cells | How to reveal hidden columns and export worksheet as PDF using Aspose.Cells .NET | C# code to unhide a range of columns and convert workbook to PDF with Aspose.Cells
// Tags: unhide columns N-P Aspose.Cells | set column width 70 points Aspose.Cells | export worksheet to PDF Aspose.Cells | column range operations Aspose.Cells C# | excel to pdf conversion after column adjustment Aspose.Cells

using System;
using Aspose.Cells;

// Loads input.xlsx, unhides columns N‑P (indexes 13‑15), sets each column width to 70 points, and saves the workbook as output.pdf using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide columns N (index 13) through P (index 15) and set their width to 70 points
        // Parameters: start column index, number of columns, width in points
        worksheet.Cells.UnhideColumns(13, 3, 70);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
