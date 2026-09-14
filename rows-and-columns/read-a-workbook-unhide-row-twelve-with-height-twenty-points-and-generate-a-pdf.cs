// Title: Unhide row 12, set its height to 20 points, and export the worksheet as a PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to make row 12 visible, assign a height of 20 points, and save the workbook as a PDF file. | Generate a PDF from an Excel workbook in C# after unhiding a specific row and adjusting its row height with Aspose.Cells.
// Common Searches: asp.net aspose.cells unhide hidden row and set row height before PDF export | c# how to change row 12 height to 20 points and convert worksheet to PDF using Aspose.Cells | unhide row and export Excel to PDF with Aspose.Cells .NET library
// Tags: unhide row Aspose.Cells C# | set row height Aspose.Cells | pdf conversion Aspose.Cells | row visibility manipulation Aspose.Cells | export worksheet to PDF C#

using System;
using Aspose.Cells;

// The example loads an existing Excel file, unhides row 12 (zero‑based index 11) and sets its height to 20 points using Cells.UnhideRow, then saves the workbook as a PDF document via Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Unhide row 12 (zero‑based index 11) and set its height to 20 points
        workbook.Worksheets[0].Cells.UnhideRow(11, 20);

        // Save the workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
