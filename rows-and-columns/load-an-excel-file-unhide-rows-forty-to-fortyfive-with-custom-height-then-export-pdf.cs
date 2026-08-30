// Title: Unhide rows 40‑45 with a custom height and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an XLSX workbook, makes rows 40‑45 visible, sets each row height to 20 points, and saves the workbook as a PDF with Aspose.Cells. | Show how to use Aspose.Cells for .NET to reveal a hidden row range, apply a specific row height, and then convert the worksheet to PDF.
// Common Searches: Aspose.Cells C# unhide hidden rows and set custom height before PDF export | Convert Excel worksheet to PDF after revealing rows with Aspose.Cells .NET | How to adjust row visibility and height then save as PDF using Aspose.Cells | C# Aspose.Cells export to PDF after modifying row dimensions
// Tags: reveal hidden rows Aspose.Cells C# | apply specific row height Aspose.Cells | export worksheet as PDF Aspose.Cells | modify row visibility Aspose.Cells .NET

using System;
using Aspose.Cells;

// Loads input.xlsx, reveals rows 40‑45 on the first worksheet, sets each row height to 20 points, and saves the workbook as output.pdf.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Path for the resulting PDF file
        string outputPath = "output.pdf";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(inputPath);

        // Unhide rows 40 to 45 (zero‑based index) and set a custom height (e.g., 20 points)
        // Total rows to unhide = 6 (rows 40,41,42,43,44,45)
        workbook.Worksheets[0].Cells.UnhideRows(40, 6, 20.0);

        // Export the workbook to PDF format
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}
