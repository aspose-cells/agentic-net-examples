// Title: Load an XLSX file, set A3 paper size on the second worksheet, and export to PDF using Aspose.Cells for .NET
// Common Searches: asp.net c# load xlsx and set A3 paper size for second sheet before PDF export | Aspose.Cells change page size of a specific worksheet to A3 and save as PDF | C# convert only the second worksheet of an XLSX to PDF with A3 layout using Aspose.Cells | how to apply A3 page setup to a particular sheet in Aspose.Cells and generate PDF | set custom paper size for a worksheet in Aspose.Cells .NET prior to PDF conversion
// Tags: worksheet A3 page setup Aspose.Cells | convert worksheet to PDF with custom size C# | Aspose.Cells set paper size for specific sheet | XLSX to PDF A3 layout Aspose.Cells | save workbook as PDF after page setup

using System;
using Aspose.Cells;

// The example loads an XLSX workbook, accesses the second worksheet, changes its PageSetup PaperSize to A3, and then saves the entire workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(inputPath);

        // Access the second worksheet (index 1, zero‑based)
        Worksheet sheet = workbook.Worksheets[1];

        // Set the paper size of the worksheet to A3
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;

        // Path for the resulting PDF file
        string outputPath = "output.pdf";

        // Save the workbook (all sheets) as a PDF document
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}
