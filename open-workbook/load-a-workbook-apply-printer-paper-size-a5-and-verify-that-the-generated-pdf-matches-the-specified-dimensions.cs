// Title: Set A5 Paper Size for a Worksheet, Export to PDF, and Verify Dimensions with Aspose.Cells (C#)
// Description: Load an Excel workbook, change the first worksheet's page setup to A5, save it as PDF, retrieve the rendered page size in inches, and confirm it matches the A5 standard within a small tolerance using Aspose.Cells.
// Keywords: Aspose.Cells A5 paper size | C# set worksheet paper size | export Excel to PDF Aspose | WorkbookRender GetPageSizeInch | validate PDF page dimensions | A5 dimensions inches | PaperSizeType.PaperA5 | SaveFormat.Pdf
// Common Searches: Aspose.Cells set worksheet to A5 | Export Excel as A5 PDF in C# | Get PDF page size from Aspose.Cells | Validate PDF dimensions programmatically | C# check A5 size after PDF conversion
// Developer Intent: Apply A5 paper size to a worksheet, generate a PDF, and programmatically verify the page dimensions.
// Use Cases: Produce printable reports that must conform to A5 size and automatically confirm the output. | Add PDF size validation to a CI/CD pipeline to ensure layout compliance. | Batch‑convert Excel files to A5 PDFs while checking each file’s page dimensions.
// AI Prompts: Generate C# code with Aspose.Cells that sets a worksheet to A5, saves it as PDF, and validates the resulting page size. | Explain how WorkbookRender.GetPageSizeInch can be used to compare PDF dimensions with the A5 standard. | Provide step‑by‑step instructions for verifying that a PDF created from an Excel workbook matches A5 dimensions within a tolerance.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an Excel workbook, change the first worksheet's page setup to A5, save it as PDF, retrieve the rendered page size in inches, and confirm it matches the A5 standard within a small tolerance using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Apply A5 paper size to the first worksheet
        workbook.Worksheets[0].PageSetup.PaperSize = PaperSizeType.PaperA5;

        // Save the workbook as PDF
        string pdfPath = "output.pdf";
        workbook.Save(pdfPath, SaveFormat.Pdf);

        // Render the workbook to obtain page dimensions in inches
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        WorkbookRender renderer = new WorkbookRender(workbook, options);
        float[] pageSize = renderer.GetPageSizeInch(0);
        double renderedWidth = Math.Round(pageSize[0], 2);
        double renderedHeight = Math.Round(pageSize[1], 2);
        Console.WriteLine($"Rendered page size (inches): {renderedWidth} x {renderedHeight}");

        // Expected A5 dimensions in inches (148 mm x 210 mm)
        double expectedWidth = Math.Round(148.0 / 25.4, 2);
        double expectedHeight = Math.Round(210.0 / 25.4, 2);
        Console.WriteLine($"Expected A5 size (inches): {expectedWidth} x {expectedHeight}");

        // Verify that the rendered size matches the expected A5 size within a small tolerance
        bool isMatch = Math.Abs(renderedWidth - expectedWidth) < 0.05 && Math.Abs(renderedHeight - expectedHeight) < 0.05;
        Console.WriteLine("Verification result: " + (isMatch ? "PASS" : "FAIL"));
    }
}
