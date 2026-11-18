// Title: How to set PDF title, author, and subject metadata before saving an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set PDF document title author subject before Save | Add custom metadata to PDF generated from Excel using Aspose.Cells .NET | Define PDF properties with PdfSaveOptions in Aspose.Cells C# example | Export Excel to PDF with metadata like Title and Author using Aspose.Cells | Set PDF document info when converting workbook to PDF in C#
// Tags: Aspose.Cells set PDF metadata C# | PdfSaveOptions DocumentInfo Aspose.Cells | export workbook to PDF with custom properties .NET | define PDF title author subject Aspose.Cells | Excel to PDF conversion metadata Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, adds sample content, ensures the output folder exists, and saves the workbook as a PDF. To embed custom metadata such as title, author, and subject, configure a PdfSaveOptions instance and set its DocumentInfo properties before invoking the Save method.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some content to the first worksheet (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample content for PDF export.");

            // Define output file path
            string outputPath = "MyWorkbook.pdf";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
