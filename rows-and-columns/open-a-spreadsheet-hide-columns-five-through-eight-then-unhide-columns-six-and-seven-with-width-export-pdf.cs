// Title: Hide columns 5‑8, unhide columns 6‑7 with custom widths, and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide a range of columns then unhide specific columns with set width | Export Excel worksheet to PDF after changing column visibility with Aspose.Cells .NET | C# code to hide columns 5 to 8 and set column 6 width to 15 and column 7 width to 20 using Aspose.Cells | How to control column visibility and width before saving workbook as PDF in Aspose.Cells for .NET | Unhide columns with custom width in Aspose.Cells and generate PDF output
// Tags: hide column range Aspose.Cells | unhide column with width Aspose.Cells | export worksheet to PDF Aspose.Cells | column visibility manipulation .NET | set column width while unhiding Aspose.Cells | Aspose.Cells PDF conversion C# | column range hide unhide Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook (or creates a new one), hides columns 5 through 8, then unhides column 6 with a width of 15 characters and column 7 with a width of 20 characters, and finally saves the worksheet as a PDF file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide columns 5 through 8 (zero‑based indices 4 to 7)
            for (int col = 4; col <= 7; col++)
            {
                sheet.Cells.HideColumn(col);
            }

            // Unhide column 6 (index 5) with width 15 characters
            sheet.Cells.UnhideColumn(5, 15);

            // Unhide column 7 (index 6) with width 20 characters
            sheet.Cells.UnhideColumn(6, 20);

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook processed successfully. PDF saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
