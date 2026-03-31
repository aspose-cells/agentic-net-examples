using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for the intermediate Excel file and the final PDF output
            string excelPath = "sample.xlsx";
            string pdfPath = "output.pdf";

            // -------------------------------------------------
            // 1. Create a workbook and populate it with sample data
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];              // get first worksheet

            // Add some data to demonstrate layout and image handling
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["C2"].PutValue(0.5);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["C3"].PutValue(0.3);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["C4"].PutValue(0.8);

            // Insert an image to illustrate image quality settings
            // (Assuming an image file "logo.png" exists in the execution directory)
            try
            {
                int pictureIndex = sheet.Pictures.Add(5, 0, "logo.png");
                sheet.Pictures[pictureIndex].Height = 100;
                sheet.Pictures[pictureIndex].Width = 100;
            }
            catch
            {
                // If the image is missing, continue without it
            }

            // Save the workbook as an intermediate Excel file (required for ConversionUtility)
            workbook.Save(excelPath);

            // -------------------------------------------------
            // 2. Configure PDF save options to control layout and image quality
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();      // create PDF save options

            // Page layout: fit all columns on one page and keep each sheet on a single page
            pdfOptions.OnePagePerSheet = true;
            pdfOptions.AllColumnsInOnePagePerSheet = true;

            // Optimize for smaller file size (image quality trade‑off)
            pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

            // Resample images to 150 DPI and set JPEG quality to 80 (higher quality)
            pdfOptions.SetImageResample(150, 80);

            // Optional: set PDF/A compliance if needed
            // pdfOptions.Compliance = PdfCompliance.PdfA1b;

            // -------------------------------------------------
            // 3. Convert the Excel file to PDF using ConversionUtility with the configured options
            // -------------------------------------------------
            LoadOptions loadOpts = new LoadOptions(LoadFormat.Xlsx);   // load as XLSX
            ConversionUtility.Convert(excelPath, loadOpts, pdfPath, pdfOptions);

            Console.WriteLine("Excel file has been successfully converted to PDF with custom layout and image quality settings.");
        }
    }
}