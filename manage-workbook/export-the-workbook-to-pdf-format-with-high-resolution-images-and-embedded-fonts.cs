using System;
using Aspose.Cells;               // Core Aspose.Cells namespace
using Aspose.Cells.Rendering;    // Required for PdfSaveOptions methods

namespace AsposeCellsPdfExport
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook (or load an existing one)
            // -------------------------------------------------
            Workbook workbook = new Workbook(); // creates a blank workbook

            // Add some sample data to demonstrate the export
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF Export");
            sheet.Cells["A2"].PutValue("High‑resolution images and embedded fonts");
            sheet.Cells["B1"].PutValue(DateTime.Now);
            // Insert an image to test high‑resolution handling
            // (Assuming an image file "sample.png" exists in the executable folder)
            try
            {
                int pictureIndex = sheet.Pictures.Add(2, 0, "sample.png");
                sheet.Pictures[pictureIndex].Height = 200; // enlarge to test resampling
                sheet.Pictures[pictureIndex].Width = 200;
            }
            catch
            {
                // If the image file is missing, continue without it
            }

            // -------------------------------------------------
            // 2. Configure PDF save options
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Embed standard Windows fonts (TrueType) into the PDF
            pdfOptions.EmbedStandardWindowsFonts = true;

            // Ensure that the workbook's default font is used when a cell's font is missing
            pdfOptions.CheckWorkbookDefaultFont = true;
            pdfOptions.DefaultFont = "Arial"; // fallback font for Unicode characters

            // Set high‑resolution image rendering (e.g., 300 DPI) and JPEG quality (90%)
            pdfOptions.SetImageResample(300, 90);

            // Optional: keep each worksheet on a single page (helps with layout)
            pdfOptions.OnePagePerSheet = true;

            // -------------------------------------------------
            // 3. Save the workbook as PDF using the Save(string, SaveOptions) overload
            // -------------------------------------------------
            string outputPath = "ExportedWorkbook.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully exported to PDF: {outputPath}");
        }
    }
}