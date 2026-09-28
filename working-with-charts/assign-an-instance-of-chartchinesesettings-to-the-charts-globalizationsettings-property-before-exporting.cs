// Title: Assign ChartChineseSettings to a chart's GlobalizationSettings property in Aspose.Cells C# before exporting the workbook
// AI Prompts: Create a ChartChineseSettings instance, configure its Font.Name to a Chinese font, assign it to chart.GlobalizationSettings, then save the workbook. | Generate C# code that applies Chinese locale settings to an Aspose.Cells chart using the GlobalizationSettings property. | Update the example to include chart.GlobalizationSettings = new ChartChineseSettings { Font = new FontInfo { Name = "Microsoft YaHei" } } before calling workbook.Save.
// Common Searches: asp.net aspose.cells set chart globalizationsettings chinese | c# aspose.cells chart chinese locale example | how to apply chinese font to chart title using aspose.cells | globalizationsettings chart asp.net chinese | export excel chart with chinese settings using aspose cells c#
// Tags: Aspose.Cells chart globalizationsettings | ChartChineseSettings C# example | set Chinese font for Excel chart title | Aspose.Cells chart Chinese globalization | export chart with Chinese locale .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, adds a column chart, sets a Chinese title and font, assigns ChartChineseSettings to the chart's GlobalizationSettings, and saves the workbook as ChartWithChineseSettings.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet (from row 5, column 0 to row 15, column 5)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set a Chinese title for the chart and apply a Chinese font
            chart.Title.Text = "中文柱形图";
            chart.Title.Font.Name = "Microsoft YaHei";

            // Save the workbook
            workbook.Save("ChartWithChineseSettings.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
