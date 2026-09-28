// Title: Generate a high‑resolution PNG of a pie chart with a custom “Other” slice using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a pie chart containing an 'Other' category, enables data labels, and saves the chart as a 300 DPI PNG with Aspose.Cells. | Adjust the sample to change the DPI to 600 and resize the chart before exporting it to a PNG file. | Enhance the example to apply a unique fill color to the 'Other' slice and include its value in the data label, then render the chart to a high‑resolution PNG.
// Common Searches: Aspose.Cells C# export pie chart with custom 'Other' slice to PNG at 300 DPI | how to set a data label for a specific point in an Aspose.Cells pie chart | render high resolution chart image using ImageOrPrintOptions in Aspose.Cells .NET | C# example creating a pie chart with an 'Other' category and saving it as PNG
// Tags: pie chart PNG export Aspose.Cells | custom slice label Aspose.Cells | high DPI chart rendering .NET | ImageOrPrintOptions resolution Aspose.Cells | C# generate workbook pie chart

using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The program creates a workbook, inserts sample data including an 'Other' category, builds a pie chart with data labels, configures ImageOrPrintOptions for 300 DPI PNG output, and saves the rendered chart as PieChartOther.png using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("Apples");
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("Bananas");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("Cherries");
            cells["B4"].PutValue(10);
            cells["A5"].PutValue("Dates");
            cells["B5"].PutValue(5);
            cells["A6"].PutValue("Other");          // Custom “Other” slice
            cells["B6"].PutValue(35);               // Sum of small categories

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Enable data labels to show category names and values
            chart.NSeries[0].DataLabels.ShowCategoryName = true;
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Note: Customizing a single point's label (e.g., “Other”) is omitted
            // because the required ChartPoint.DataLabel property is not available
            // in the current Aspose.Cells version.

            // Configure high‑resolution PNG output (300 DPI)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300
                // ImageFormat defaults to PNG; other properties omitted for compatibility
            };

            // Determine output path and ensure directory exists
            string outputFile = "PieChartOther.png";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Render the chart to a PNG file
            chart.ToImage(outputFile, imgOptions);
            Console.WriteLine($"Chart image saved to: {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
