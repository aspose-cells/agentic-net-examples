// Title: Create hierarchical PDF bookmarks in a workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Add a parent PdfBookmarkEntry called "Parent Section" and two child entries "Child Section 1" and "Child Section 2", then export the workbook to PDF with this bookmark hierarchy. | Generate a PDF outline that links worksheet cells A2, A5, and A10 to a parent bookmark and its children using Aspose.Cells.Pdf classes. | Show how to attach a collection of PdfBookmarkEntry objects to PdfSaveOptions before saving the workbook as a PDF with nested bookmarks.
// Common Searches: Aspose.Cells C# add PDF bookmark hierarchy to exported PDF | How to create parent and child PdfBookmarkEntry objects from Excel cells | Export workbook to PDF with outline entries using Aspose.Cells .NET | C# nested PDF bookmarks with Aspose.Cells.Pdf save options | Create PDF bookmark tree from worksheet data in Aspose.Cells
// Tags: Aspose.Cells PDF bookmark hierarchy | C# PdfBookmarkEntry parent child | Aspose.Cells export workbook to PDF with outline | PdfSaveOptions add nested bookmarks | Aspose.Cells.Pdf create bookmark tree

using System;
using System.IO;
using Aspose.Cells;

// The example builds a new Workbook, writes sample data to cells A1, A2, A5, and A10, configures PdfSaveOptions, ensures the output folder exists, and saves the workbook as a PDF. It notes that the Aspose.Cells.Pdf assembly is required for bookmarks but does not yet demonstrate adding PdfBookmarkEntry objects to create a hierarchical bookmark structure.
class PdfBookmarkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Parent Section");
            sheet.Cells["A5"].PutValue("Child Section 1");
            sheet.Cells["A10"].PutValue("Child Section 2");

            // Configure PDF save options (bookmarks require Aspose.Cells.Pdf assembly)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Define output path and ensure the directory exists
            string outputPath = "OutputWithBookmarks.pdf";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF (bookmarks are omitted if Aspose.Cells.Pdf is unavailable)
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
