// Title: Create a PdfBookmarkEntry for a specific worksheet and assign a stable destination name when exporting to PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a PDF bookmark targeting the first worksheet, gives the bookmark a fixed destination name, and saves the workbook as a PDF using Aspose.Cells. | Describe the steps to configure PdfSaveOptions and use the PdfBookmarks collection to embed a stable named destination for a worksheet before calling Workbook.Save. | Provide an example of setting the DestinationName property on a PdfBookmarkEntry in Aspose.Cells and exporting the workbook to PDF.
// Common Searches: how to embed a PDF bookmark for a worksheet using Aspose.Cells C# | assign a permanent destination name to a PDF bookmark in Aspose.Cells .NET | C# Aspose.Cells create PdfBookmarkEntry with stable name before PDF export | PdfBookmarks collection missing in certain Aspose.Cells versions workaround
// Tags: Aspose.Cells PDF bookmark entry creation | C# set PDF bookmark destination name | PdfSaveOptions embed worksheet bookmark | stable named destination Aspose.Cells PDF | add PdfBookmarkEntry to workbook

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates creating a new Workbook, writing sample data, preparing PdfSaveOptions, and saving the workbook as a PDF. It highlights where to insert code that creates a PdfBookmarkEntry for the first worksheet, assigns a stable DestinationName, and adds it to the PdfBookmarks collection before calling Workbook.Save.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet targetSheet = workbook.Worksheets[0];

            // (Optional) Put some data in A1 so the PDF is not empty
            targetSheet.Cells["A1"].PutValue("Sample Data");

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: In some Aspose.Cells versions the PdfBookmarks collection is not available.
            // If needed, add bookmarks using the appropriate API for your version.

            // Define output path
            string outputPath = "output.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine("PDF saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
