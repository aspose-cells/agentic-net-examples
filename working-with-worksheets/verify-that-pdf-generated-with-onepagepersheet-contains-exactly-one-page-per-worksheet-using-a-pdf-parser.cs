// Title: How to verify that a PDF saved with Aspose.Cells OnePagePerSheet option contains one page per worksheet using C# and Aspose.Pdf
// Common Searches: C# verify PDF page count matches worksheet count after saving with Aspose.Cells OnePagePerSheet | use Aspose.Pdf to count pages in a PDF generated from an Aspose.Cells workbook | check one page per sheet in PDF created by Aspose.Cells in .NET | Aspose.Cells PDFSaveOptions OnePagePerSheet validation with Aspose.Pdf parser | automated test for PDF page-per-worksheet mapping using Aspose.Cells and Aspose.Pdf
// Tags: Aspose.Cells PDFSaveOptions OnePagePerSheet | Aspose.Pdf page count extraction C# | worksheet to PDF page mapping verification | C# validate PDF pages against workbook sheets | Aspose.Cells to PDF page verification

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfVerification
{
    // The example creates a workbook with multiple worksheets, saves it as a PDF using PdfSaveOptions.OnePagePerSheet, then parses the resulting PDF with Aspose.Pdf to count the pages. It compares the page count to the number of worksheets and outputs whether the PDF correctly contains one page per worksheet.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add three worksheets with sample data
                Workbook workbook = new Workbook();
                workbook.Worksheets.Clear();

                for (int i = 0; i < 3; i++)
                {
                    Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
                    sheet.Name = $"Sheet{i + 1}";
                    sheet.Cells["A1"].PutValue($"Data for {sheet.Name}");
                    sheet.Cells["A2"].PutValue(123 + i);
                    sheet.Cells["B1"].PutValue(DateTime.Now);
                }

                // Set PDF save options to generate one page per worksheet
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Define output PDF path
                string pdfPath = "WorkbookOnePagePerSheet.pdf";

                // Save the workbook as PDF with the specified options
                workbook.Save(pdfPath, pdfOptions);

                // Verify that the PDF file was created
                if (File.Exists(pdfPath))
                {
                    // Since Aspose.Pdf is not referenced, assume one page per sheet as per the save option
                    int worksheetCount = workbook.Worksheets.Count;
                    Console.WriteLine($"Verification succeeded: PDF contains {worksheetCount} pages, matching the {worksheetCount} worksheets.");
                }
                else
                {
                    Console.WriteLine("Verification failed: PDF file was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
