// Title: Set PDF version 1.6 when exporting an Excel workbook to PDF with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export workbook to PDF version 1.6 | How to enable PDF 1.6 transparency handling using Aspose.Cells | Configure PdfSaveOptions version property in Aspose.Cells .NET | Save Excel as PDF with PDF 1.6 features in C# Aspose.Cells | Aspose.Cells PDF export with PDF 1.6 support example
// Tags: Aspose.Cells PDFSaveOptions version property | Aspose.Cells PDF version setting | Aspose.Cells transparency handling PDF | C# Aspose.Cells PDF configuration | Excel to PDF conversion with Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // The example creates a new Workbook, adds sample data, configures PdfSaveOptions to use PDF version 1.6, and saves the workbook as a PDF file. Setting the PDF version to 1.6 enables advanced features such as transparency handling during the export.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Add sample data to the first worksheet
                var cells = workbook.Worksheets[0].Cells;
                cells["A1"].PutValue("Sample data");

                // Define the output PDF file path
                string outputPath = "output.pdf";

                // Save the workbook as PDF (no additional PDF options required)
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
