// Title: Unhide column B, set width to 50 points, and export worksheet to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, reveal column B, set its width to 50 points, and save the worksheet as a PDF using Aspose.Cells in C#. | Using Aspose.Cells for .NET, unhide column B, assign a 50‑point width, and export the workbook to PDF. | In C#, programmatically display column B, set column width to 50 points, then convert the first worksheet to a PDF with Aspose.Cells.
// Common Searches: how to make hidden columns visible and set width in Aspose.Cells before PDF conversion | adjust column width in points with Aspose.Cells and generate PDF output | Aspose.Cells .NET export Excel to PDF after column formatting
// Tags: column B visibility Aspose.Cells | column width 50 points Aspose.Cells | PDF export of worksheet Aspose.Cells | Excel column formatting C# Aspose.Cells | convert Excel to PDF Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads 'input.xlsx', unhides column B (index 1) and sets its width to 50 points, then saves the workbook as 'output.pdf' using the PDF save format.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide column B (zero‑based index 1) and set its width to 50 points
        worksheet.Cells.UnhideColumn(1, 50);

        // Save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
