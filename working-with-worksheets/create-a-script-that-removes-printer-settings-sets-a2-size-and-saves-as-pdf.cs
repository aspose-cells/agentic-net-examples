// Title: Remove worksheet printer settings, set A2 paper size, and save as PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, clears the PrinterSettings for every worksheet, changes the PaperSize to A2, and saves the result as a PDF. | Show how to iterate through all worksheets, set PageSetup.PrinterSettings to null, apply PaperSizeType.PaperA2, and export the workbook to a PDF file using Aspose.Cells.
// Common Searches: aspnet remove printer settings from Excel worksheets before PDF conversion | set A2 page size for all sheets using Aspose.Cells C# | export Excel to PDF with custom page setup Aspose.Cells .NET | how to clear PageSetup.PrinterSettings in Aspose.Cells
// Tags: clear worksheet printer settings Aspose.Cells | set A2 paper size Aspose.Cells | convert workbook to PDF Aspose.Cells | modify page setup all worksheets .NET | printer settings nullification Aspose.Cells

using System;
using Aspose.Cells;

// // Loads an Excel file, clears the PrinterSettings for each worksheet, sets each sheet's PaperSize to A2, and saves the workbook as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your source file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets to modify page setup
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup setup = sheet.PageSetup;

            // Remove any printer settings associated with the worksheet
            setup.PrinterSettings = null;

            // Set the paper size to A2
            setup.PaperSize = PaperSizeType.PaperA2;
        }

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
