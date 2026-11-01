// Title: Add a custom title to an Excel chart and export it as a PNG image with Aspose.Cells for .NET
// Common Searches: how to set a chart title in Aspose.Cells before saving as PNG in C# | Aspose.Cells C# export chart to PNG with custom caption | render Excel chart to image with title using Aspose.Cells .NET | C# code to add visible title to chart and generate PNG file with Aspose.Cells | Aspose.Cells chart.Title property example for PNG output
// Tags: Aspose.Cells chart title styling | C# generate PNG from Excel chart Aspose.Cells | chart.Title visibility control .NET | custom chart font settings Aspose.Cells | render Excel chart image programmatically

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads a workbook, accesses the first chart, assigns a custom visible title with specific font settings, and then renders the chart directly to a PNG file using Aspose.Cells for .NET.
class ChartCaptionExample
{
    static void Main()
    {
        try
        {
            const string inputFile = "InputWithChart.xlsx";
            const string outputFile = "ChartWithCaption.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"The input file '{inputFile}' was not found.");

            // Load an existing workbook that contains a chart
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet and its first chart
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts found in the first worksheet.");

            Chart chart = sheet.Charts[0];

            // Add a caption as the chart title
            string caption = "Sales Performance Q1 2024";
            chart.Title.Text = caption;
            chart.Title.IsVisible = true;
            chart.Title.Font.Size = 14;
            chart.Title.Font.IsBold = true;
            chart.Title.Font.Color = Color.Black; // Requires System.Drawing.Common at runtime

            // Render the chart directly to a PNG file
            chart.ToImage(outputFile); // PNG is the default format

            Console.WriteLine($"Chart image with caption saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
