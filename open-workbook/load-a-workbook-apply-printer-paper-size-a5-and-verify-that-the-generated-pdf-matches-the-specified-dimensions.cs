// Title: C# – Load Excel with Aspose.Cells, set A5 paper size, export to PDF, and validate dimensions
// Description: Demonstrates how to load an Excel workbook using Aspose.Cells LoadOptions, apply the A5 paper size (PaperSizeType.PaperA5), save the workbook as a PDF, retrieve the rendered page size with WorkbookRender.GetPageSizeInch, and confirm that the PDF dimensions (5.83" × 8.27") fall within a 0.05‑inch tolerance.
// Keywords: Aspose.Cells C# A5 paper size | LoadOptions SetPaperSize | WorkbookRender GetPageSizeInch | save workbook as PDF Aspose.Cells | verify PDF page dimensions | paper size validation Aspose.Cells | Excel to PDF A5
// Common Searches: Aspose.Cells set A5 paper size when loading workbook | How to check PDF page size after exporting Excel with Aspose.Cells | Get page dimensions in inches from WorkbookRender | Validate PDF dimensions against A5 using Aspose.Cells .NET
// Developer Intent: Apply A5 paper size to a workbook during load, convert it to PDF, and programmatically ensure the resulting PDF matches A5 dimensions.
// Use Cases: Create mobile‑friendly reports that must fit A5 sheets before distribution. | Automate quality control in a document‑generation pipeline to confirm PDF page size after conversion. | Batch‑process invoices or tickets where each PDF must conform to A5 envelope specifications.
// AI Prompts: Generate C# code with Aspose.Cells that loads an Excel file, sets A5 paper size via LoadOptions, saves it as PDF, and verifies the page size using WorkbookRender. | Explain how WorkbookRender.GetPageSizeInch returns width and height in inches and how to compare them to standard A5 measurements. | Suggest strategies for handling unit conversion and tolerance adjustments when validating PDF dimensions with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPaperSizeDemo
{
    // Demonstrates how to load an Excel workbook using Aspose.Cells LoadOptions, apply the A5 paper size (PaperSizeType.PaperA5), save the workbook as a PDF, retrieve the rendered page size with WorkbookRender.GetPageSizeInch, and confirm that the PDF dimensions (5.83" × 8.27") fall within a 0.05‑inch tolerance.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path) with A5 paper size set via LoadOptions
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.SetPaperSize(PaperSizeType.PaperA5);
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Verify that the first worksheet inherits the A5 setting
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Worksheet PageSetup PaperSize: " + sheet.PageSetup.PaperSize);

            // Save the workbook as PDF – the PDF will use the A5 paper size defined above
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine("Workbook saved as PDF to: " + pdfPath);

            // Use WorkbookRender to obtain the actual page dimensions in inches
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
            float[] pageSizeInch = renderer.GetPageSizeInch(0); // [0] = width, [1] = height

            // A5 size in inches (rounded to two decimal places)
            const float A5WidthInch = 5.83f;  // 148 mm
            const float A5HeightInch = 8.27f; // 210 mm

            Console.WriteLine($"Rendered page size: {pageSizeInch[0]:0.00}\" x {pageSizeInch[1]:0.00}\"");

            // Simple verification against expected A5 dimensions (allowing a small tolerance)
            const float tolerance = 0.05f; // 0.05 inch tolerance
            bool widthMatches = Math.Abs(pageSizeInch[0] - A5WidthInch) <= tolerance;
            bool heightMatches = Math.Abs(pageSizeInch[1] - A5HeightInch) <= tolerance;

            if (widthMatches && heightMatches)
                Console.WriteLine("Verification passed: PDF page size matches A5 dimensions.");
            else
                Console.WriteLine("Verification failed: PDF page size does not match A5 dimensions.");

            // Clean up
            renderer.Dispose();
        }
    }
}
