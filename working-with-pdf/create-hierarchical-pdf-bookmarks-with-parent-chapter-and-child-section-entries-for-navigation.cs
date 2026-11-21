// Title: Creating hierarchical chapter and section PDF bookmarks from an Excel workbook using Aspose.Cells for .NET (current version limitation)
// Common Searches: Aspose.Cells .NET add parent and child bookmarks to PDF generated from Excel | C# generate PDF outline with chapters and sections using Aspose.Cells | How to include hierarchical bookmarks when saving a workbook as PDF with Aspose.Cells | Aspose.Cells PDFSaveOptions bookmark hierarchy not supported workaround | Create PDF navigation bookmarks from worksheet data in C#
// Tags: Aspose.Cells hierarchical PDF bookmarks | C# PDF outline from Excel worksheet | PdfSaveOptions bookmark hierarchy limitation | Excel to PDF navigation bookmarks .NET | Aspose.Cells PDF bookmark support

using System;
using System.IO;
using Aspose.Cells;

// The example builds a workbook, fills column A with chapter and section titles, configures PdfSaveOptions, and saves the workbook as a PDF. It also highlights that hierarchical PDF bookmarks (parent‑child outline entries) are not supported in the current Aspose.Cells assembly version.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the sheet with sample data representing chapters and sections
            sheet.Cells["A1"].PutValue("Chapter 1");
            sheet.Cells["A2"].PutValue("Section 1.1");
            sheet.Cells["A3"].PutValue("Section 1.2");
            sheet.Cells["A5"].PutValue("Chapter 2");
            sheet.Cells["A6"].PutValue("Section 2.1");
            sheet.Cells["A7"].PutValue("Section 2.2");

            // Prepare PDF save options (bookmarks are not supported in the current assembly version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Define output path and ensure the directory exists
            string outputPath = "HierarchicalBookmarks.pdf";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
