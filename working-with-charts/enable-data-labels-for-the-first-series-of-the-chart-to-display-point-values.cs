// Title: How to enable data labels that display point values for the first series of a column chart in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to turn on ShowValue for the first series of a column chart so each column shows its numeric value. | Add a step to an existing Aspose.Cells workbook that activates data labels for the first series of a chart and then saves the file. | Demonstrate how to calculate a chart after enabling data labels for its first series and export the workbook as an .xlsx file in C#. | Show the minimal Aspose.Cells API calls required to set DataLabels.ShowValue = true on the first series of a chart.
// Common Searches: Aspose.Cells C# enable data labels for first series of a column chart | How to show values on each column in an Aspose.Cells chart using .NET | Set ShowValue property on chart series with Aspose.Cells API | C# example to display point values in a column chart created by Aspose.Cells | Enable data labels for a specific series in an Aspose.Cells workbook
// Tags: Aspose.Cells enable data labels first series | column chart series ShowValue .NET | C# chart point values Aspose.Cells | set data labels for chart series Aspose.Cells | calculate and save chart Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // // This example creates a workbook, fills it with sample data, adds a column chart, enables data labels for the first series to display each point's value, calculates the chart, and saves the workbook as an .xlsx file.
    public class EnableDataLabelsFirstSeries
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series to show point values
            Series firstSeries = chart.NSeries[0];
            firstSeries.DataLabels.ShowValue = true;

            // Calculate the chart (optional)
            chart.Calculate();

            // Save the workbook
            workbook.Save("EnableDataLabelsFirstSeries.xlsx");
        }
    }
}
