// Title: Set Letter paper size with LoadOptions and verify PDF page count using Aspose.Cells for .NET
// Description: Demonstrates how to create LoadOptions, set the default printer paper size to Letter, load an .xlsx workbook, render it with ImageOrPrintOptions (OnePagePerSheet), retrieve the rendered page count via WorkbookRender, compare it to an expected value, and optionally save the result as a PDF.
// Keywords: Aspose.Cells | C# | .NET | LoadOptions | SetPaperSize | PaperLetter | PDF page count | WorkbookRender | OnePagePerSheet | printer paper size | Excel to PDF conversion
// Common Searches: Aspose.Cells set paper size programmatically | LoadOptions SetPaperSize Letter example | how to get PDF page count from Aspose.Cells | verify rendered page count Aspose.Cells .NET | render Excel to PDF with specific paper size
// Developer Intent: Apply a Letter paper size when loading a workbook and confirm that the generated PDF contains the expected number of pages.
// Use Cases: Ensure consistent print layout across environments by loading Excel files with a predefined Letter paper size. | Automate PDF generation in CI/CD pipelines, asserting the output page count for regression testing. | Create unit tests that validate workbook rendering produces the correct number of PDF pages before distribution.
// AI Prompts: Generate C# code that loads an Excel workbook with LoadOptions.SetPaperSize(PaperSizeType.PaperLetter) and returns the PDF page count using WorkbookRender. | Write a C# unit test that opens a workbook, sets Letter paper size, renders to PDF, and asserts the page count equals a given value. | Explain the effect of ImageOrPrintOptions.OnePagePerSheet on page count calculation when converting an Excel workbook to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPaperSizeDemo
{
    // Demonstrates how to create LoadOptions, set the default printer paper size to Letter, load an .xlsx workbook, render it with ImageOrPrintOptions (OnePagePerSheet), retrieve the rendered page count via WorkbookRender, compare it to an expected value, and optionally save the result as a PDF.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (replace with an actual file path)
            string sourcePath = "input.xlsx";

            // Expected number of PDF pages after rendering (set according to your test case)
            int expectedPageCount = 1;

            // Create LoadOptions and set the default printer paper size to Letter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.SetPaperSize(PaperSizeType.PaperLetter);

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Optional: verify that the worksheet's page setup reflects the Letter size
            Console.WriteLine("Worksheet PaperSize after load: " +
                workbook.Worksheets[0].PageSetup.PaperSize);

            // Create print options for rendering
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions
            {
                // Ensure each sheet is rendered as a separate page (helps predict page count)
                OnePagePerSheet = true
            };

            // Render the workbook to evaluate the total page count
            WorkbookRender renderer = new WorkbookRender(workbook, printOptions);
            int actualPageCount = renderer.PageCount;

            Console.WriteLine($"Evaluated PDF page count: {actualPageCount}");

            // Verify the page count matches the expectation
            if (actualPageCount == expectedPageCount)
            {
                Console.WriteLine("Page count verification succeeded.");
            }
            else
            {
                Console.WriteLine($"Page count verification failed. Expected {expectedPageCount}, but got {actualPageCount}.");
            }

            // Save the workbook as PDF (optional, demonstrates the final output)
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to: {pdfPath}");
        }
    }
}
