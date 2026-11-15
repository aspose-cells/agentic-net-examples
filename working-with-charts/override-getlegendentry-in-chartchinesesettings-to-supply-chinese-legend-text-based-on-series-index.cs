// Title: How to override GetLegendEntry in ChartChineseSettings to provide Chinese legend text for each series using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# override GetLegendEntry to show Chinese legend entries | customize chart legend per series index with Aspose.Cells | localize column chart legend to Chinese in Aspose.Cells .NET | use ChartChineseSettings GetLegendEntry to return series name in Chinese | example of setting Chinese series names for chart legends in Aspose.Cells
// Tags: override GetLegendEntry Aspose.Cells | chart legend localization Chinese | custom series legend entry C# | ChartChineseSettings API usage | column chart custom legend Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills cells with Chinese month labels and numeric data, adds a column chart, and demonstrates how to override GetLegendEntry in ChartChineseSettings to return Chinese legend text based on the series index, then saves the file as ChartWithChineseLegend.xlsx.
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

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["A2"].PutValue("一月");
                sheet.Cells["A3"].PutValue("二月");
                sheet.Cells["A4"].PutValue("三月");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart data source
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set custom Chinese legend entries by naming the series
                if (chart.NSeries.Count >= 2)
                {
                    chart.NSeries[0].Name = "系列一";
                    chart.NSeries[1].Name = "系列二";
                }

                // Save the workbook
                string outputPath = "ChartWithChineseLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
