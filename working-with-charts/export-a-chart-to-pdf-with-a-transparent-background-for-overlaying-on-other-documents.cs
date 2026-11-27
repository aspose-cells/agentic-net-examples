// Title: Export an Aspose.Cells chart to a PDF with a transparent background using C#
// Common Searches: C# render Aspose.Cells chart as transparent PNG and embed in PDF | Aspose.Cells export chart with transparent background to PDF in .NET | How to save a chart image with transparency and place it in a PDF using Aspose.Cells | Create PDF overlay with a transparent chart image using Aspose.Cells C# | Aspose.Cells example for chart to PDF with transparent background
// Tags: chart rendering transparent PNG Aspose.Cells | embed picture shape into PDF workbook Aspose.Cells | ImageOrPrintOptions Transparent property C# | PdfSaveOptions OnePagePerSheet Aspose.Cells | free floating picture placement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;   // Needed for PlacementType and Picture

// Shows how to create a column chart, render it to a transparent PNG using ImageOrPrintOptions, insert the PNG as a free‑floating picture into a new workbook, and save the workbook as a PDF with PdfSaveOptions.
class Program
{
    static void Main()
    {
        // Paths (adjust as needed)
        string chartImagePath = "chart_transparent.png";
        string outputPdfPath = "ChartOverlay.pdf";

        try
        {
            // ------------------------------------------------------------
            // 1. Create a workbook and add a sample chart
            // ------------------------------------------------------------
            Workbook sourceWb = new Workbook();
            Worksheet ws = sourceWb.Worksheets[0];

            // Populate some data for the chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Sample Chart";

            // ------------------------------------------------------------
            // 2. Render the chart to a PNG image with a transparent background
            // ------------------------------------------------------------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                Transparent = true,               // makes background transparent
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Save the chart as a transparent PNG
            chart.ToImage(chartImagePath, imgOptions);
            Console.WriteLine($"Chart image saved to '{chartImagePath}' with transparent background.");

            // Ensure the image file was created
            if (!File.Exists(chartImagePath))
                throw new FileNotFoundException($"Chart image not found: {chartImagePath}");

            // ------------------------------------------------------------
            // 3. Create a new workbook that will be saved as PDF
            // ------------------------------------------------------------
            Workbook pdfWb = new Workbook();
            Worksheet pdfWs = pdfWb.Worksheets[0];

            // Insert the transparent PNG as a picture shape
            int pictureIdx = pdfWs.Pictures.Add(0, 0, chartImagePath);
            Picture picture = pdfWs.Pictures[pictureIdx];
            picture.Placement = PlacementType.FreeFloating; // Allows exact positioning
            picture.Left = 0;
            picture.Top = 0;

            // Optional: set print area to the top‑left cell
            pdfWs.PageSetup.PrintArea = "A1";

            // ------------------------------------------------------------
            // 4. Save the workbook as PDF
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            pdfWb.Save(outputPdfPath, pdfOptions);
            Console.WriteLine($"PDF saved to '{outputPdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
