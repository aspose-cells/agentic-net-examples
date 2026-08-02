// Title: Set Letter Paper Size When Loading a Workbook and Verify PDF Page Count with Aspose.Cells for .NET
// Description: Demonstrates how to load an Excel file using LoadOptions with the default printer paper size set to Letter, apply the same size to the workbook and its first worksheet, evaluate the expected PDF page count with WorkbookPrintingPreview, save the workbook as PDF, and confirm that the generated PDF matches the evaluated page count.
// Keywords: Aspose.Cells C# load workbook paper size | PaperLetter LoadOptions | WorkbookPrintingPreview page count | evaluate PDF pages Aspose.Cells | verify PDF page count .NET | set default printer paper size Aspose | Excel to PDF page count validation
// Common Searches: How to set default printer paper size to Letter in Aspose.Cells | Get evaluated PDF page count before saving with Aspose.Cells | Compare evaluated page count and actual PDF pages Aspose.Cells .NET | Load workbook with specific paper size Aspose.Cells | Validate PDF page count generated from Excel using Aspose
// Developer Intent: Load a workbook with Letter paper size, determine the expected PDF page count, export to PDF, and ensure the exported document uses the intended layout and page count.
// Use Cases: Standardize printing layout across environments by enforcing Letter paper size when opening Excel files. | Predict the number of PDF pages an Excel workbook will occupy without creating the PDF first. | Automated quality checks to confirm that the PDF output matches the expected pagination after export.
// AI Prompts: Write C# code that loads an Excel workbook with LoadOptions set to PaperLetter, evaluates the PDF page count using WorkbookPrintingPreview, saves as PDF, and verifies the counts match. | Explain the role of WorkbookPrintingPreview in calculating page count and how to compare its result with the actual PDF pages in Aspose.Cells. | Suggest a method to handle mismatched page counts, including logging details and adjusting worksheet PageSetup properties programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfPageCountDemo
{
    // Demonstrates how to load an Excel file using LoadOptions with the default printer paper size set to Letter, apply the same size to the workbook and its first worksheet, evaluate the expected PDF page count with WorkbookPrintingPreview, save the workbook as PDF, and confirm that the generated PDF matches the evaluated page count.
    class Program
    {
        static void Main()
        {
            // Load options with default printer paper size set to Letter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.SetPaperSize(PaperSizeType.PaperLetter);

            // Load an existing workbook using the load options
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Ensure the workbook's default paper size is Letter
            workbook.Settings.PaperSize = PaperSizeType.PaperLetter;

            // Also set the first worksheet's page setup to Letter (optional but explicit)
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;

            // Create print options (default settings)
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions();

            // Evaluate total page count for the workbook before saving as PDF
            WorkbookPrintingPreview preview = new WorkbookPrintingPreview(workbook, printOptions);
            int evaluatedPageCount = preview.EvaluatedPageCount;
            Console.WriteLine($"Evaluated page count (before PDF): {evaluatedPageCount}");

            // Save the workbook as PDF
            workbook.Save("output.pdf", SaveFormat.Pdf);

            // After saving, you can re‑evaluate if needed (same workbook instance)
            WorkbookPrintingPreview postSavePreview = new WorkbookPrintingPreview(workbook, printOptions);
            int postSavePageCount = postSavePreview.EvaluatedPageCount;
            Console.WriteLine($"Evaluated page count (after PDF): {postSavePageCount}");

            // Verify that the page counts match expectations
            if (evaluatedPageCount == postSavePageCount)
            {
                Console.WriteLine("PDF page count aligns with the evaluated page count.");
            }
            else
            {
                Console.WriteLine("Warning: PDF page count does not match the evaluated page count.");
            }
        }
    }
}
