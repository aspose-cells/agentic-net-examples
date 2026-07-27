// Title: Set Letter paper size via LoadOptions and verify PDF page count with AspNet Aspose.Cells
// Description: Demonstrates how to create LoadOptions, apply PaperSizeType.PaperLetter, load an Excel workbook, configure ImageOrPrintOptions for PDF, use WorkbookRender to read the page count before saving, and compare the actual count with an expected value in C#.
// Keywords: Aspose.Cells C# | LoadOptions SetPaperSize | PaperLetter | WorkbookRender PageCount | Excel to PDF conversion | validate PDF pagination | OnePagePerSheet | printer paper size | PDF page count verification | automated PDF testing
// Common Searches: Aspose.Cells set default printer paper size to Letter | Get PDF page count before saving with Aspose.Cells | How to verify Excel to PDF page count in .NET | LoadOptions SetPaperSize example | WorkbookRender page count usage
// Developer Intent: Load an Excel file with Letter paper dimensions, render it to PDF, retrieve the generated page count, and confirm it matches a predefined expectation.
// Use Cases: Standardize printed layout across environments by forcing Letter paper size during import. | Pre‑emptively determine PDF pagination to drive conditional workflow logic. | Integrate PDF page‑count validation into CI/CD pipelines for Excel‑to‑PDF conversion jobs.
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions to set PaperSizeType.PaperLetter, loads an Excel workbook, renders it to PDF with one page per sheet, obtains the page count via WorkbookRender, and asserts the count equals a given number. | Explain the effect of LoadOptions.SetPaperSize on PDF output in Aspose.Cells and describe how WorkbookRender.PageCount can be used for pagination checks before saving.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create LoadOptions, apply PaperSizeType.PaperLetter, load an Excel workbook, configure ImageOrPrintOptions for PDF, use WorkbookRender to read the page count before saving, and compare the actual count with an expected value in C#.
class SetPaperSizeAndVerifyPdfPageCount
{
    static void Main()
    {
        // Path to the source Excel file
        string sourceFile = "input.xlsx";

        // Create LoadOptions and set the default printer paper size to Letter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.SetPaperSize(PaperSizeType.PaperLetter);

        // Load the workbook with the specified LoadOptions
        Workbook workbook = new Workbook(sourceFile, loadOptions);

        // Prepare print options for PDF rendering
        ImageOrPrintOptions printOptions = new ImageOrPrintOptions
        {
            SaveFormat = SaveFormat.Pdf,
            OnePagePerSheet = true // ensure each sheet is rendered as a separate page
        };

        // Use WorkbookRender to evaluate the total page count before saving
        WorkbookRender renderer = new WorkbookRender(workbook, printOptions);
        int actualPageCount = renderer.PageCount;

        // Save the workbook as PDF
        string pdfFile = "output.pdf";
        workbook.Save(pdfFile, SaveFormat.Pdf);

        // Define the expected page count (adjust as needed for your test case)
        int expectedPageCount = 1;

        // Verify that the actual page count matches the expectation
        if (actualPageCount == expectedPageCount)
        {
            Console.WriteLine($"Success: PDF page count ({actualPageCount}) matches expected value.");
        }
        else
        {
            Console.WriteLine($"Failure: PDF page count ({actualPageCount}) does not match expected ({expectedPageCount}).");
        }

        // Clean up resources
        renderer.Dispose();
    }
}
