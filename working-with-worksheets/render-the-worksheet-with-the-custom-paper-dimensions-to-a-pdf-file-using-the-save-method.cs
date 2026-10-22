// Title: Export a worksheet with custom paper size and fit-to-page settings to PDF using Aspose.Cells in C#
// AI Prompts: Generate C# code that configures a worksheet's PaperSize, Orientation, FitToPagesWide, and FitToPagesTall properties with Aspose.Cells and then saves the workbook as a PDF. | Modify the sample to apply a non‑standard paper size (e.g., 8 in × 10 in) before exporting the worksheet to PDF using Aspose.Cells.
// Common Searches: Aspose.Cells C# set custom paper dimensions for worksheet before PDF export | How to fit worksheet content to one page wide when saving as PDF using Aspose.Cells | C# example for changing page orientation and paper size then converting workbook to PDF with Aspose.Cells
// Tags: worksheet page setup with Aspose.Cells | fit-to-page PDF conversion Aspose.Cells | set nonstandard paper dimensions Aspose.Cells | configure page orientation Aspose.Cells C# | convert workbook to PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program creates a new workbook, accesses the first worksheet, sets the page setup to Letter size, portrait orientation, and fits the content to one page wide, then saves the workbook as a PDF file named CustomPaperWorksheet.pdf.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set paper size to standard Letter (8.5 x 11 inches)
            sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;

            // Set orientation and fit settings
            sheet.PageSetup.Orientation = PageOrientationType.Portrait;
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;

            // Define output PDF path
            string outputPath = "CustomPaperWorksheet.pdf";

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
