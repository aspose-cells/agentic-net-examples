// Title: How to compare PDF file sizes when using AllColumnsInOnePagePerSheet versus default export with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# compare PDF file size with AllColumnsInOnePagePerSheet option | measure difference between default PDF export and AllColumnsInOnePagePerSheet in .NET | C# code to get size of PDFs generated from Excel using PdfSaveOptions
// Tags: AllColumnsInOnePagePerSheet PDF export size impact | PdfSaveOptions default configuration size effect | Aspose.Cells Excel to PDF file size analysis | C# PDF generation size measurement using Aspose.Cells | Aspose.Cells PDF export performance metrics

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, saves it twice as PDF—once with default PdfSaveOptions and once with AllColumnsInOnePagePerSheet enabled—then reads the resulting file sizes, prints both sizes, and indicates whether the AllColumnsInOnePage PDF is larger, smaller, or equal to the default PDF, with exception handling.
class PdfSizeComparison
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Paths for the generated PDF files
            string defaultPdfPath = "default.pdf";
            string allColumnsPdfPath = "AllColumnsInOnePage.pdf";

            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(excelPath);

            // -------------------------------------------------
            // Export with default PDF settings
            // -------------------------------------------------
            PdfSaveOptions defaultOptions = new PdfSaveOptions();
            workbook.Save(defaultPdfPath, defaultOptions);

            // -------------------------------------------------
            // Export with AllColumnsInOnePagePerSheet enabled
            // -------------------------------------------------
            PdfSaveOptions allColumnsOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = true
            };
            workbook.Save(allColumnsPdfPath, allColumnsOptions);

            // -------------------------------------------------
            // Compare file sizes
            // -------------------------------------------------
            long defaultSize = new FileInfo(defaultPdfPath).Length;
            long allColumnsSize = new FileInfo(allColumnsPdfPath).Length;

            Console.WriteLine($"Default PDF size: {defaultSize} bytes");
            Console.WriteLine($"AllColumnsInOnePage PDF size: {allColumnsSize} bytes");

            if (allColumnsSize > defaultSize)
            {
                Console.WriteLine("The AllColumnsInOnePage PDF is larger than the default PDF.");
            }
            else if (allColumnsSize < defaultSize)
            {
                Console.WriteLine("The AllColumnsInOnePage PDF is smaller than the default PDF.");
            }
            else
            {
                Console.WriteLine("Both PDFs have the same file size.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
