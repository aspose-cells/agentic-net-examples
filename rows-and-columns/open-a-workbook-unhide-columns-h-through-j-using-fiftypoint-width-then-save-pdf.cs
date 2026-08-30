// Title: How to unhide columns H‑J, set a 50‑point width, and export an Excel workbook to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel file, unhide columns H through J, set each column width to 50 points, and save the workbook as a PDF with Aspose.Cells in C#. | Using Aspose.Cells, reveal hidden columns H‑J, apply a 50‑point column width, then convert the worksheet to a PDF document.
// Common Searches: Aspose.Cells C# unhide columns H J and set column width before PDF export | Set column width in points with Aspose.Cells and convert workbook to PDF | How to reveal hidden columns and adjust width in an Excel file using Aspose.Cells .NET | Convert Excel to PDF after modifying column visibility with Aspose.Cells C#
// Tags: column unhide operation Aspose.Cells | column width setting points Aspose.Cells | PDF generation from workbook Aspose.Cells | worksheet column visibility Aspose.Cells | adjust column dimensions before PDF conversion Aspose.Cells

using Aspose.Cells;
using System;

// // Loads 'input.xlsx', unhides columns H through J (indices 7‑9) with a width of 50 points each, and saves the workbook as 'output.pdf' in PDF format using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Unhide columns H (index 7) through J (index 9) and set their width to 50 points
        cells.UnhideColumns(7, 3, 50.0);

        // Save the modified workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
