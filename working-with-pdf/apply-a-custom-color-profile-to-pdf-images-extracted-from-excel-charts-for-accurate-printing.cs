using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering; // for PdfSaveOptions and PdfCompliance

class ApplyCustomColorProfile
{
    static void Main()
    {
        // Load the source workbook that contains the chart
        string sourceWorkbookPath = "source.xlsx";
        Workbook sourceWorkbook = new Workbook(sourceWorkbookPath);

        // Access the first worksheet and its first chart
        Worksheet sourceWorksheet = sourceWorkbook.Worksheets[0];
        if (sourceWorksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the source workbook.");
            return;
        }
        Chart chart = sourceWorksheet.Charts[0];

        // Render the chart to a PNG image (PNG preserves colour profile information)
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png // ensures the image is saved as PNG
        };

        using (MemoryStream chartImageStream = new MemoryStream())
        {
            // Convert chart to image and write to the memory stream
            chart.ToImage(chartImageStream, imgOptions);
            chartImageStream.Position = 0; // reset stream position for later use

            // Create a new workbook that will be saved as PDF
            Workbook pdfWorkbook = new Workbook();
            Worksheet pdfWorksheet = pdfWorkbook.Worksheets[0];

            // Insert the chart image into the worksheet
            // The picture will be placed at cell A1 (row 0, column 0)
            pdfWorksheet.Pictures.Add(0, 0, chartImageStream);

            // Configure PDF save options for high‑quality printing
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use PDF/A‑1b compliance which is suitable for archival/printing
                Compliance = PdfCompliance.PdfA1b
            };

            // Save the workbook (containing the PNG chart image) as PDF
            string outputPdfPath = "ChartWithCustomColorProfile.pdf";
            pdfWorkbook.Save(outputPdfPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to '{outputPdfPath}'.");
        }
    }
}