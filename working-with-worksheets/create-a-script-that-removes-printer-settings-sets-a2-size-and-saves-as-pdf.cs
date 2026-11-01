// Title: C# – Clear worksheet printer settings, set paper size to A2, and export Excel to PDF with Aspose.Cells
// Common Searches: how to clear printer settings of all worksheets and set A2 paper size before converting to PDF with Aspose.Cells in C# | Aspose.Cells C# remove print area and printer settings then save workbook as PDF | set worksheet PaperSize to A2 using Aspose.Cells and export to PDF | reset PrinterSettings property for each sheet in Aspose.Cells before PDF generation | clear print titles and centering options in Aspose.Cells C#
// Tags: worksheet printer configuration reset Aspose.Cells | A2 paper size setting Aspose.Cells C# | PDF export with custom page setup Aspose.Cells | remove print area and titles Aspose.Cells | nullify PrinterSettings object Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook, iterates through every worksheet, clears printer‑related properties (PrintArea, PrintTitleRows, PrintTitleColumns, centering flags, and the PrinterSettings object), sets the paper size to A2, and then saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Process each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;

            // Remove printer settings by clearing related properties
            pageSetup.PrintArea = "";
            pageSetup.PrintTitleRows = "";
            pageSetup.PrintTitleColumns = "";
            pageSetup.CenterHorizontally = false;
            pageSetup.CenterVertically = false;

            // If a PrinterSettings object exists, reset it
            pageSetup.PrinterSettings = null;

            // Set the paper size to A2
            pageSetup.PaperSize = PaperSizeType.PaperA2;
        }

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
