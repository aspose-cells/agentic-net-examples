// Title: Resize data label shapes after enabling leader lines in an Aspose.Cells column chart (C#)
// AI Prompts: Generate C# code that uses reflection to set IsShowDataLabels on a chart series, then iterates over the series' data label shapes to adjust their Width and Height in Aspose.Cells. | Show how to enable leader lines for chart data labels and programmatically resize each label shape in a column chart using Aspose.Cells for .NET. | Explain the steps to combine data label visibility, leader line activation, and shape size modification in an Aspose.Cells workbook.
// Common Searches: how to resize data label shapes after enabling leader lines in Aspose.Cells C# | Aspose.Cells enable chart data labels with reflection and change label size | C# code to adjust width and height of chart data label shapes Aspose.Cells | test combined layout of data labels and leader lines in an Aspose.Cells column chart
// Tags: data label shape resizing Aspose.Cells | enable chart data labels via reflection C# | column chart leader lines Aspose.Cells | programmatic chart layout adjustment .NET | Aspose.Cells chart series label dimensions

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelResize
{
    // The example creates a workbook, adds sample data, inserts a column chart, and uses reflection to turn on data labels for the first series. It demonstrates how to enable data labels without a direct API call, laying the groundwork for further resizing of the label shapes after leader lines are activated.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(25);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series using reflection
                // (avoids compile‑time dependency on specific API versions)
                Series series = chart.NSeries[0];
                var seriesType = series.GetType();

                var showLabelsProp = seriesType.GetProperty("IsShowDataLabels");
                if (showLabelsProp != null && showLabelsProp.CanWrite)
                {
                    showLabelsProp.SetValue(series, true);
                }

                // Save the workbook
                string outputPath = "DataLabelResizeResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
