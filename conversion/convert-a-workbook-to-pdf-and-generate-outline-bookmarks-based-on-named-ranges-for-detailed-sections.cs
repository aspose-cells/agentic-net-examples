// Title: Create a PDF from an Excel workbook with outline bookmarks generated from named ranges using Aspose.Cells for .NET (C#)
// Common Searches: c# aspose.cells convert excel to pdf with outline bookmarks from named ranges | how to generate pdf bookmarks from excel named ranges using aspose.cells | aspose.cells pdfsaveoptions preserve named range hierarchy in c# | export workbook to pdf each sheet on a separate page aspose.cells | fit all columns on one pdf page per sheet using aspose.cells
// Tags: Aspose.Cells PDFSaveOptions outline bookmarks | Excel named ranges to PDF bookmarks .NET | C# workbook to PDF one page per sheet | Aspose.Cells fit all columns one page per sheet | Aspose.Cells preserve named range hierarchy in PDF

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the source Excel file, loads it with Aspose.Cells, configures PdfSaveOptions to fit all columns on a single page per sheet and start each worksheet on a new PDF page, then saves the workbook as a PDF where defined named ranges are automatically converted into outline bookmarks. Errors are caught and reported.
class WorkbookToPdfWithBookmarks
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Optional: fit all columns on one page per sheet
                AllColumnsInOnePagePerSheet = true,

                // Optional: each worksheet will start on a new PDF page
                OnePagePerSheet = true
            };

            // Save the workbook as PDF; named ranges will become outline bookmarks automatically
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully created at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
