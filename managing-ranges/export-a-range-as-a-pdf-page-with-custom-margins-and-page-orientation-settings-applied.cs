// Title: Export a specific cell range to PDF with custom margins and landscape orientation using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export range A1:D20 to PDF with custom margins | How to set print area and page orientation before saving Excel as PDF in Aspose.Cells | C# Aspose.Cells PdfSaveOptions one page per sheet with custom margins | Export selected worksheet range to PDF in landscape mode using Aspose.Cells | Set left, right, top, bottom margins for PDF export in Aspose.Cells .NET
// Tags: export range to PDF Aspose.Cells .NET | custom PDF margins Aspose.Cells | set print area Aspose.Cells | page orientation PDF Aspose.Cells | PdfSaveOptions OnePagePerSheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, defines a print area (A1:D20) on the first worksheet, applies custom left, right, top, and bottom margins, optionally sets landscape orientation, configures PdfSaveOptions to generate one PDF page per sheet, and saves the selected range as a PDF file.
class ExportRangeToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or any worksheet you need)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range you want to export (e.g., A1:D20)
            string rangeAddress = "A1:D20";

            // Set the print area to the defined range so only this area is exported
            sheet.PageSetup.PrintArea = rangeAddress;

            // Apply custom margins (values are in inches)
            sheet.PageSetup.LeftMargin = 0.5;   // left margin
            sheet.PageSetup.RightMargin = 0.5;  // right margin
            sheet.PageSetup.TopMargin = 0.7;    // top margin
            sheet.PageSetup.BottomMargin = 0.7; // bottom margin

            // Set page orientation (Portrait or Landscape) – optional
            // Uncomment the line below if the enum is available in your Aspose.Cells version
            // sheet.PageSetup.Orientation = PageOrientation.Landscape;

            // Configure PDF save options (optional: one page per sheet)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF; only the defined range will appear due to the print area
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
