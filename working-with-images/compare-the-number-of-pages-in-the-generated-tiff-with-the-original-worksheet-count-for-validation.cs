// Title: C# code to verify that a multi‑page TIFF created from an Excel workbook has the same number of pages as worksheets using Aspose.Cells
// Common Searches: how to get page count of a TIFF generated from Excel using Aspose.Cells C# | compare number of worksheets to TIFF pages after rendering with Aspose.Cells | validate multi‑page TIFF page count matches Excel sheet count in .NET | use WorkbookPrintingPreview evaluated page count to check TIFF pages in C# | Aspose.Cells render workbook to multi‑page TIFF and verify page count
// Tags: Aspose.Cells render workbook to multi‑page TIFF | WorkbookPrintingPreview evaluate TIFF page count | compare worksheet count with TIFF pages C# | multi‑page TIFF validation Aspose.Cells .NET | ImageOrPrintOptions TIFF rendering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffValidation
{
    // The sample loads an Excel file (or creates a three‑sheet workbook), renders it to a multi‑page TIFF using ImageOrPrintOptions, obtains the generated TIFF's page count via WorkbookPrintingPreview.EvaluatedPageCount, and compares this count to the workbook's worksheet count, outputting whether the validation succeeded.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input Excel file path
            string excelPath = args.Length > 0 ? args[0] : "input.xlsx";
            string tiffPath = Path.ChangeExtension(excelPath, ".tiff");

            // Ensure the Excel file exists; if not, create a sample workbook with multiple worksheets
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Excel file \"{excelPath}\" not found. Creating a sample workbook.");
                CreateSampleWorkbook(excelPath);
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(excelPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook \"{excelPath}\": {ex.Message}");
                return;
            }

            // Get the original worksheet count
            int worksheetCount = workbook.Worksheets.Count;
            Console.WriteLine($"Original worksheet count: {worksheetCount}");

            // Set up image/print options for TIFF rendering
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Tiff,
                TiffCompression = TiffCompression.CompressionLZW,
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Render the workbook to a multi‑page TIFF
            try
            {
                WorkbookRender renderer = new WorkbookRender(workbook, options);
                renderer.ToImage(tiffPath);
                Console.WriteLine($"Workbook rendered to TIFF: {tiffPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to render TIFF: {ex.Message}");
                return;
            }

            // Verify that the TIFF file was created
            if (!File.Exists(tiffPath))
            {
                Console.WriteLine($"TIFF file \"{tiffPath}\" was not created.");
                return;
            }

            // Determine the number of pages in the generated TIFF using Aspose.Cells preview API
            int tiffPageCount;
            try
            {
                WorkbookPrintingPreview preview = new WorkbookPrintingPreview(workbook, options);
                tiffPageCount = preview.EvaluatedPageCount;
                Console.WriteLine($"Generated TIFF page count (via preview): {tiffPageCount}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to evaluate TIFF page count: {ex.Message}");
                return;
            }

            // Compare the counts and output validation result
            if (tiffPageCount == worksheetCount)
            {
                Console.WriteLine("Validation succeeded: TIFF page count matches worksheet count.");
            }
            else
            {
                Console.WriteLine($"Validation failed: TIFF page count ({tiffPageCount}) does not match worksheet count ({worksheetCount}).");
            }
        }

        /// <param name="path">File path where the workbook will be saved.</param>
        private static void CreateSampleWorkbook(string path)
        {
            Workbook wb = new Workbook();

            // Worksheet 1 (default)
            Worksheet ws1 = wb.Worksheets[0];
            ws1.Name = "Sheet1";
            ws1.Cells["A1"].PutValue("Sample data in Sheet1");

            // Worksheet 2
            int idx2 = wb.Worksheets.Add();
            Worksheet ws2 = wb.Worksheets[idx2];
            ws2.Name = "Sheet2";
            ws2.Cells["A1"].PutValue("Sample data in Sheet2");

            // Worksheet 3
            int idx3 = wb.Worksheets.Add();
            Worksheet ws3 = wb.Worksheets[idx3];
            ws3.Name = "Sheet3";
            ws3.Cells["A1"].PutValue("Sample data in Sheet3");

            try
            {
                wb.Save(path, SaveFormat.Xlsx);
                Console.WriteLine($"Sample workbook saved to \"{path}\" with {wb.Worksheets.Count} worksheets.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save sample workbook: {ex.Message}");
                throw;
            }
        }
    }
}
