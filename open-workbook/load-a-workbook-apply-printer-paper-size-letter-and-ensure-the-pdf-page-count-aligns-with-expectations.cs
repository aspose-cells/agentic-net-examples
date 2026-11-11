// Title: Load an Excel workbook, set each worksheet to Letter paper size, export to PDF, and validate the PDF page count with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set paper size to Letter for all worksheets before PDF export | How to get page count of PDF generated from Excel using WorkbookRender in .NET | Validate expected number of pages after converting XLSX to PDF with Aspose.Cells | C# apply printer settings to every sheet and check resulting PDF page count | Aspose.Cells save workbook as PDF and compare actual page count to expected value
// Tags: letter paper size worksheet Aspose.Cells | pdf export custom paper size Aspose.Cells | workbookrender pdf page count .NET | pdf page count verification C# | apply printer settings all sheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel file, sets the Letter paper size on all worksheets, saves the workbook as a PDF, then uses WorkbookRender with PDF options to retrieve the actual page count and compares it to an expected value, outputting the verification result.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string pdfPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Apply Letter paper size to all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }

            // Save the workbook as a PDF file
            workbook.Save(pdfPath, SaveFormat.Pdf);

            // Use WorkbookRender to obtain the page count of the generated PDF
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Pdf
            };
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
            int actualPageCount = renderer.PageCount;

            // Define the expected page count (adjust as needed)
            int expectedPageCount = 2;

            // Verify the page count
            if (actualPageCount == expectedPageCount)
            {
                Console.WriteLine($"PDF page count verification passed: {actualPageCount} pages.");
            }
            else
            {
                Console.WriteLine($"PDF page count verification failed: expected {expectedPageCount}, but got {actualPageCount}.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
