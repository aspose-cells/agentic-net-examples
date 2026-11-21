// Title: Set Japanese globalization settings for an Excel chart using Aspose.Cells for .NET before exporting
// Common Searches: Aspose.Cells C# assign ChartJapaneseSettings to chart GlobalizationSettings | how to apply Japanese localization to an Excel chart with Aspose.Cells before saving | export chart with Japanese title using Aspose.Cells globalization property in .NET | C# example for setting chart GlobalizationSettings to Japanese in Aspose.Cells | apply Japanese number and date formats to a chart in Aspose.Cells workbook
// Tags: ChartJapaneseSettings Aspose.Cells | chart globalization settings C# | Japanese chart localization Aspose.Cells | apply globalization to Excel chart .NET | export chart with Japanese title Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// This example creates a workbook, adds a column chart with sample data and a Japanese title, then creates a ChartJapaneseSettings instance and assigns it to the chart's GlobalizationSettings property before saving, ensuring the chart is correctly localized for Japanese users.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

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

            // Optional: set a Japanese title for the chart
            chart.Title.Text = "サンプルチャート";

            // Save the workbook
            workbook.Save("ChartWithJapaneseSettings.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
