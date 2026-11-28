// Title: Export an Excel chart to JPEG with a custom branding background color using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export chart to JPEG with specific background color | set chart area background color before saving chart as image in Aspose.Cells | convert Excel chart to high‑resolution JPEG using ImageOrPrintOptions in .NET | apply brand color to chart area and plot area in Aspose.Cells | save workbook chart as JPEG file with 300 DPI using Aspose.Cells
// Tags: chart area foreground color Aspose.Cells | plot area background color Aspose.Cells | ImageOrPrintOptions high DPI JPEG Aspose.Cells | convert workbook chart to image C# | branding color chart export Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example loads a workbook (creating a sample one if needed), accesses the first chart, applies a branding RGB color to both the chart area and plot area, configures JPEG output options (90 % quality, 300 dpi), ensures the output folder exists, and saves the chart as a JPEG image using Chart.ToImage.
class Program
{
    static void Main()
    {
        // Paths – adjust as needed
        string inputWorkbookPath = "input.xlsx";
        string outputChartImagePath = "chart_output.jpg";

        try
        {
            // Ensure the input workbook exists; if not, create a sample workbook with a chart
            Workbook workbook;
            if (File.Exists(inputWorkbookPath))
            {
                workbook = new Workbook(inputWorkbookPath);
            }
            else
            {
                workbook = CreateSampleWorkbookWithChart();
                workbook.Save(inputWorkbookPath);
                Console.WriteLine($"Sample workbook created at '{inputWorkbookPath}'.");
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that a chart exists
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // ---------- Apply branding background color ----------
            // Example branding color – replace with your exact brand RGB values
            Color brandingColor = Color.FromArgb(30, 144, 255); // DodgerBlue

            // Set the chart area background
            chart.ChartArea.Area.ForegroundColor = brandingColor;

            // Optionally also set the plot area background to keep a uniform look
            chart.PlotArea.Area.ForegroundColor = brandingColor;

            // ---------- Configure JPEG output options ----------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Jpeg,
                Quality = 90,                 // JPEG quality (0‑100)
                HorizontalResolution = 300,   // DPI
                VerticalResolution = 300      // DPI
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputChartImagePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // ---------- Convert chart to JPEG ----------
            try
            {
                chart.ToImage(outputChartImagePath, imgOptions);
                Console.WriteLine($"Chart successfully saved as JPEG to: {outputChartImagePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save chart image: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Creates a simple workbook with sample data and a column chart
    private static Workbook CreateSampleWorkbookWithChart()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Sample data
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["A2"].PutValue("Apple");
        ws.Cells["A3"].PutValue("Orange");
        ws.Cells["A4"].PutValue("Banana");

        ws.Cells["B1"].PutValue("Sales");
        ws.Cells["B2"].PutValue(1200);
        ws.Cells["B3"].PutValue(800);
        ws.Cells["B4"].PutValue(1500);

        // Add a column chart
        int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = ws.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);

        return wb;
    }
}
