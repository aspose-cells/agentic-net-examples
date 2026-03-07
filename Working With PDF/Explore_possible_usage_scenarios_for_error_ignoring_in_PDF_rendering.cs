using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsErrorIgnoringDemo
{
    public class Program
    {
        public static void Main()
        {
            // 1. Hide all rendering errors (shapes, images, charts) using PdfSaveOptions.IgnoreError
            IgnoreRenderingErrorsDemo();

            // 2. Ignore formula calculation errors before PDF conversion
            CalculationIgnoreErrorDemo();

            // 3. Control how cell errors are displayed in the PDF using PageSetup.PrintErrors
            PrintErrorDisplayDemo();

            // 4. Capture rendering warnings via WarningCallback and decide to ignore specific ones
            WarningCallbackDemo();
        }

        private static void IgnoreRenderingErrorsDemo()
        {
            // Create a workbook with a shape that references a missing image (will cause a rendering error)
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Normal text");

            // Use an empty stream to simulate a missing image
            using (MemoryStream emptyImageStream = new MemoryStream())
            {
                ws.Shapes.AddPicture(5, 5, 100, 100, emptyImageStream);
            }

            // Set IgnoreError to true to suppress rendering errors
            PdfSaveOptions opts = new PdfSaveOptions
            {
                IgnoreError = true
            };

            // Save to PDF; any shape/image/chart errors are hidden
            wb.Save("IgnoreRenderingErrors.pdf", opts);
        }

        private static void CalculationIgnoreErrorDemo()
        {
            // Create a workbook with a formula that uses an unsupported function
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].Formula = "=UNKNOWNFUNC(1)";

            // Calculate formulas while ignoring calculation errors
            wb.CalculateFormula(new CalculationOptions { IgnoreError = true });

            // Save to PDF; the cell will contain the error value but no exception is thrown
            PdfSaveOptions opts = new PdfSaveOptions();
            wb.Save("CalculationIgnoreError.pdf", opts);
        }

        private static void PrintErrorDisplayDemo()
        {
            // Create a workbook with a formula that results in a division error
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].Formula = "=1/0";

            // Display errors as dashes in the rendered PDF
            ws.PageSetup.PrintErrors = PrintErrorsType.PrintErrorsDash;

            // Save to PDF; the error appears as "--" instead of "#DIV/0!"
            PdfSaveOptions opts = new PdfSaveOptions();
            wb.Save("PrintErrorDash.pdf", opts);
        }

        private static void WarningCallbackDemo()
        {
            // Create a workbook with a cell that uses a non‑existent font (triggers a warning)
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Text with missing font");
            Style style = wb.CreateStyle();
            style.Font.Name = "NonExistentFont";
            ws.Cells["A1"].SetStyle(style);

            // Assign a custom warning callback to capture rendering warnings
            PdfSaveOptions opts = new PdfSaveOptions
            {
                WarningCallback = new RenderingWarningCallback()
            };

            // Save to PDF; warnings are logged but rendering continues
            wb.Save("WarningCallbackDemo.pdf", opts);
        }

        // Custom implementation of IWarningCallback to handle rendering warnings
        private class RenderingWarningCallback : IWarningCallback
        {
            public void Warning(WarningInfo warningInfo)
            {
                // Log the warning type and description
                Console.WriteLine($"Warning: {warningInfo.Type} - {warningInfo.Description}");

                // Example: silently ignore font substitution warnings
                // No exception is thrown, allowing rendering to proceed
            }
        }
    }
}