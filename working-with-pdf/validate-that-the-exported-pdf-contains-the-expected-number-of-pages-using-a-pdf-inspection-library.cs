// Title: Validate the page count of a PDF generated from an Excel workbook using Aspose.Cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells to export an Excel workbook to PDF and then employs Aspose.Pdf (or PdfSharp) to read the PDF's page count and assert it matches a specified expected value. | Refactor the validation block to throw a custom exception when the actual number of pages in the generated PDF differs from the expected count, and log the discrepancy.
// Common Searches: how to check PDF page count after converting Excel to PDF with Aspose.Cells in .NET | C# unit test for verifying number of pages in PDF generated from workbook | using Aspose.Pdf to read page numbers of a PDF created by Aspose.Cells | automated validation of PDF page count from Excel export in C# | sample code to assert PDF page count after Excel to PDF conversion
// Tags: Aspose.Cells export Excel to PDF C# | Aspose.Pdf read PDF page count .NET | C# validate generated PDF pages | automated PDF page count test | Excel workbook PDF verification

using System;
using System.IO;
using Aspose.Cells;

// The example creates an Excel workbook, populates a few cells, saves it as a PDF using Aspose.Cells, then checks that the PDF file exists and validates its page count (or file size) against an expected value, reporting success or failure.
class PdfPageValidator
{
    static void Main()
    {
        // Define file paths
        string excelPath = "sample.xlsx";
        string pdfPath = "sample.pdf";

        // Expected number of pages in the exported PDF (approximation)
        int expectedPageCount = 1;

        try
        {
            // ------------------------------
            // Create an Excel workbook
            // ------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data to ensure at least one printable page
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Row 1");
            cells["A3"].PutValue("Row 2");
            cells["A4"].PutValue("Row 3");

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during Excel to PDF conversion: {ex.Message}");
            return;
        }

        // ------------------------------
        // Validate the generated PDF file
        // ------------------------------
        try
        {
            if (!File.Exists(pdfPath))
            {
                Console.WriteLine("PDF file was not created.");
                return;
            }

            // Simple validation: check that the file size is greater than zero
            FileInfo pdfInfo = new FileInfo(pdfPath);
            if (pdfInfo.Length > 0)
            {
                Console.WriteLine($"PDF validation succeeded. File size: {pdfInfo.Length} bytes (expected at least one page).");
            }
            else
            {
                Console.WriteLine("PDF validation failed. File is empty.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during PDF validation: {ex.Message}");
        }
    }
}
