// Title: Load an Excel workbook, set all worksheets to Letter paper size, export to PDF, and verify the PDF page count with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set Letter paper size for every worksheet before PDF conversion | How to retrieve PDF page count after converting Excel to PDF using WorkbookRender | Configure printing settings with LoadOptions when loading a workbook in Aspose.Cells | Validate expected number of PDF pages generated from an Excel file in .NET | C# example to compare actual PDF page count with a predefined value after export
// Tags: letter paper size setting for worksheets Aspose.Cells | load workbook with LoadOptions C# | excel to pdf page count validation Aspose.Cells | WorkbookRender retrieve page count | printing options configuration prior to pdf export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for WorkbookRender

// This C# example demonstrates loading an Excel file using LoadOptions, applying the Letter paper size to each worksheet, saving the workbook as a PDF, obtaining the generated PDF's page count via WorkbookRender, and comparing it against an expected value to confirm successful conversion.
class Program
{
    static void Main()
    {
        try
        {
            // Expected number of PDF pages after conversion (adjust as needed)
            int expectedPageCount = 2;

            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook (LoadOptions can be customized if needed)
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Ensure each worksheet uses Letter paper size (8.5 x 11 inches)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }

            // Save the workbook to PDF (temporary file)
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);

            // Obtain the page count of the generated PDF using WorkbookRender
            int actualPageCount;
            try
            {
                // ImageOrPrintOptions is required for WorkbookRender constructor
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();
                WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
                actualPageCount = renderer.PageCount;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to render workbook: {ex.Message}");
                return;
            }

            // Verify that the PDF page count matches the expectation
            if (actualPageCount == expectedPageCount)
            {
                Console.WriteLine($"Success: PDF contains the expected {expectedPageCount} pages.");
            }
            else
            {
                Console.WriteLine($"Failure: PDF contains {actualPageCount} pages, expected {expectedPageCount}.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
