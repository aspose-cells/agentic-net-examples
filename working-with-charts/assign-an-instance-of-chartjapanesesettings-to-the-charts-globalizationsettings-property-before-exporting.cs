// Title: Assign ChartJapaneseSettings to a chart’s GlobalizationSettings property in Aspose.Cells C# before exporting the workbook
// AI Prompts: Generate C# code that creates a ChartJapaneseSettings object for the "ja-JP" culture and assigns it to chart.GlobalizationSettings in Aspose.Cells before saving the workbook. | Show how to configure Japanese number and date formats on an Aspose.Cells chart by setting its GlobalizationSettings with ChartJapaneseSettings. | Write a snippet that demonstrates applying ChartJapaneseSettings to a column chart and then exporting the workbook to an .xlsx file.
// Common Searches: Aspose.Cells C# set chart globalization to Japanese locale | ChartJapaneseSettings usage example for Aspose.Cells chart | How to apply Japanese culture to a chart in Aspose.Cells before saving | Assign chart-level globalization settings in Aspose.Cells .NET | Export Aspose.Cells chart with Japanese number formatting
// Tags: chart globalization settings Aspose.Cells C# | ChartJapaneseSettings assignment .NET | Japanese locale chart formatting Aspose.Cells | set chart culture Aspose.Cells | export chart with Japanese globalization Aspose.Cells

using System;
using System.Drawing.Imaging;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, sets its culture to Japanese, adds sample data, inserts a column chart, assigns a ChartJapaneseSettings instance to the chart's GlobalizationSettings property, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set Japanese culture for the workbook (globalization)
            workbook.Settings.CultureInfo = new CultureInfo("ja-JP");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data series for the chart
            chart.NSeries.Add("A1:A3", true);

            // Optional: Export the chart as an image (requires System.Drawing.Common)
            // Uncomment the following line if the required assembly is referenced.
            // chart.ToImage("chart.png", ImageFormat.Png);

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
