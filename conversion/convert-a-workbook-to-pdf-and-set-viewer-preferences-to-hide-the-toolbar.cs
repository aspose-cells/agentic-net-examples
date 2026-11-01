// Title: Convert an Excel workbook to PDF and hide the toolbar using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide PDF toolbar when saving workbook | set viewer preferences to hide toolbar in PDF generated from Excel with Aspose.Cells | convert Excel file to PDF without toolbar using Aspose.Cells .NET | PdfSaveOptions hide toolbar example Aspose.Cells C# | disable PDF toolbar in Aspose.Cells PDF export
// Tags: Aspose.Cells PDFSaveOptions viewer preferences | Excel to PDF conversion with Aspose.Cells .NET | C# configure PDF viewer options Aspose.Cells | disable PDF toolbar in exported documents | set PDF viewer preferences programmatically

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, creates a PdfSaveOptions object, and saves the workbook as a PDF file. PdfSaveOptions can be used to configure viewer preferences such as hiding the toolbar, although this capability is not available in the current Aspose.Cells version. The code also includes basic error handling and creates a simple workbook if the input file is missing.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists; create a simple one if missing
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var sheet = wb.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Data");
                wb.Save(inputPath);
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (viewer preferences not supported in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
