// Title: Unhide row 12, set its height to 20 points, and save the worksheet as a PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook, makes row 12 visible, changes its height to 20 points, and saves the file as a PDF using Aspose.Cells. | Show how to adjust row visibility and row height in Aspose.Cells before converting a worksheet to PDF in a .NET application. | Provide a step‑by‑step example for un‑hiding a specific row, setting its height, and exporting the sheet to PDF with Aspose.Cells for C#.
// Common Searches: Aspose.Cells C# unhide specific row and set row height before PDF conversion | How to change row 12 visibility and height in an Excel file and export to PDF using Aspose.Cells .NET | C# code to unhide hidden row, set height to 20 points, and save workbook as PDF with Aspose.Cells
// Tags: row visibility manipulation Aspose.Cells C# | set specific row height Aspose.Cells PDF export | zero‑based row indexing Aspose.Cells | convert worksheet to PDF Aspose.Cells C# | modify worksheet rows before PDF generation Aspose.Cells

using System;
using Aspose.Cells;

// // Loads input.xlsx, unhides row 12 (index 11), sets its height to 20 points, and saves the workbook as output.pdf in PDF format using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Row numbers are zero‑based; row 12 is index 11
        Row row12 = sheet.Cells.Rows[11];

        // Unhide the row
        row12.IsHidden = false;

        // Set the row height to 20 points
        row12.Height = 20;

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
