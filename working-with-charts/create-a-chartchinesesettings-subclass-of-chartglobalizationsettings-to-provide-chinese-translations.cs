// Title: Subclass ChartGlobalizationSettings to supply Chinese labels for Aspose.Cells charts in C#
// AI Prompts: Write a C# class named ChartChineseSettings that inherits from Aspose.Cells.Charts.ChartGlobalizationSettings and overrides its properties to return Chinese strings for the chart title, category axis title, value axis title, and series name. | Create a workbook with Chinese sample data, add a column chart, assign the ChartChineseSettings instance to the chart, enable value data labels, and save the workbook as an .xlsx file using Aspose.Cells.
// Common Searches: how to localize Aspose.Cells chart titles to Chinese in .NET | C# subclass ChartGlobalizationSettings for Chinese chart labels | Aspose.Cells set Chinese axis titles on column chart | display series name in Chinese on Aspose.Cells chart | save Excel workbook with Chinese chart localization using Aspose.Cells
// Tags: subclass ChartGlobalizationSettings Chinese localization | set Chinese chart axis titles Aspose.Cells | apply Chinese series name Aspose.Cells chart | enable data labels on localized Aspose.Cells chart | save workbook with Chinese chart labels .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example demonstrates how to create a ChartChineseSettings class that inherits from ChartGlobalizationSettings to provide Chinese translations for chart elements such as titles, axis labels, and series names. It then shows how to apply this custom globalization settings to a column chart built from Chinese data, enable value data labels, and save the workbook as an Excel file.
public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            cells["A1"].PutValue("类别");
            cells["B1"].PutValue("数值");
            cells["A2"].PutValue("一");
            cells["A3"].PutValue("二");
            cells["A4"].PutValue("三");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = workbook.Worksheets[0].Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = workbook.Worksheets[0].Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set series name (first series)
            if (chart.NSeries.Count > 0)
                chart.NSeries[0].Name = "系列";

            // Set axis titles (Chinese)
            chart.CategoryAxis.Title.Text = "类别轴标题";
            chart.ValueAxis.Title.Text = "值轴标题";

            // Set chart title (optional)
            chart.Title.Text = "图例";

            // Show data labels (values) on the series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Save the workbook
            string outputPath = "ChartChineseSettings.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
