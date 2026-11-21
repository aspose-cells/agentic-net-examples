// Title: How to set PDF document title from the original Excel workbook name when converting with Aspose.Cells for .NET (unsupported)
// Common Searches: Aspose.Cells C# set PDF document title to Excel file name during conversion | Can I preserve Excel workbook name as PDF title using Aspose.Cells .NET | Aspose.Cells PDFSaveOptions title metadata not applied in C# | Set PDF metadata title property when saving workbook as PDF with Aspose.Cells | Why does Aspose.Cells fail to set PDF title metadata from source workbook
// Tags: Aspose.Cells PDFSaveOptions title metadata | C# Excel to PDF conversion preserve filename | Aspose.Cells set PDF document properties | PDF metadata limitation Aspose.Cells .NET | Workbook filename to PDF title Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook and saves it as a PDF using Aspose.Cells for .NET. It highlights that the current Aspose.Cells version does not support directly setting PDF document metadata such as the title, so the PDF title cannot be automatically derived from the original Excel file name.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(excelPath))
                {
                    throw new FileNotFoundException($"The input Excel file was not found: {excelPath}");
                }

                // Load the Excel workbook
                Workbook workbook = new Workbook(excelPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // NOTE: Setting PDF document metadata (e.g., title) is not supported directly
                // in the current Aspose.Cells version used. The conversion will still succeed.

                // Save the workbook as PDF
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Conversion successful. PDF saved to: {pdfPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
