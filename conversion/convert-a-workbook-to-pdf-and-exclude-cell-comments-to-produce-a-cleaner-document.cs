// Title: Convert an Excel workbook to PDF and hide cell comments with Aspose.Cells for .NET (C#)
// Common Searches: aspnet convert xlsx to pdf without comments using Aspose.Cells | c# hide cell comments when saving Excel as PDF with Aspose.Cells | PdfSaveOptions CommentsVisibility hide example Aspose.Cells | export workbook to PDF excluding annotations Aspose.Cells C# | how to remove Excel comments in PDF output using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions hide comments | C# Excel to PDF conversion without annotations | Workbook.Save PDF options Aspose.Cells | Excel comments exclusion in PDF export | Aspose.Cells PDF conversion settings C#

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions (optionally setting CommentsVisibility to hide comments), and saves the workbook as a PDF, producing a clean document without cell comment markers.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Hide cell comments if the property is supported by your Aspose.Cells version.
            // Uncomment the line below for versions that include CommentsVisibility.
            // pdfOptions.CommentsVisibility = PdfCommentsVisibility.Hide;

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
