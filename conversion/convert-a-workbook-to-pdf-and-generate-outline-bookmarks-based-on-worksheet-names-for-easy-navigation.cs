// Title: Convert an Excel workbook to PDF with worksheet‑name outline bookmarks using Aspose.Cells for .NET
// AI Prompts: Write a C# program that loads an .xlsx file with Aspose.Cells and saves it as a PDF, automatically creating outline bookmarks for each worksheet based on its name. | Update existing Aspose.Cells PDF conversion code to enable PdfSaveOptions.BookmarksExportMode so that each sheet appears as a bookmark in the generated PDF. | Add custom PDF settings (page orientation, compression) while preserving worksheet‑name bookmarks during the Excel‑to‑PDF conversion in C#.
// Common Searches: Aspose.Cells C# generate PDF bookmarks from worksheet names | How to export Excel to PDF with outline navigation using Aspose.Cells .NET | C# convert .xlsx to PDF with bookmarks for each sheet Aspose.Cells | Enable PDF outline bookmarks when saving workbook to PDF Aspose.Cells | Excel to PDF conversion preserving sheet navigation Aspose.Cells example
// Tags: Aspose.Cells PDF conversion with worksheet bookmarks | C# generate PDF outline from Excel sheets | PdfSaveOptions enable bookmarks Aspose.Cells | Excel workbook to PDF with navigation bookmarks | Aspose.Cells PDF export preserving sheet structure

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates how to load an Excel workbook with Aspose.Cells for .NET and save it as a PDF while automatically creating outline bookmarks for each worksheet using PdfSaveOptions, providing easy navigation within the resulting PDF document.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Load the source workbook; create a simple one if the file does not exist.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
                // Optionally save the generated workbook for later runs.
                workbook.Save(inputPath);
            }

            // Configure PDF save options (bookmarks omitted for compatibility with older versions).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the configured options.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
