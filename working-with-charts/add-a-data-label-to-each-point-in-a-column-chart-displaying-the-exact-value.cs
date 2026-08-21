// Title: Aspose.Cells C# – Show Exact Values on Each Column Chart Bar with Data Labels
// Description: Creates a workbook, fills A1:B4 with categories and numbers, inserts a column chart, binds the series, enables data labels to display each point’s value, positions them outside the columns, and saves the file as XLSX.
// Keywords: Aspose.Cells | C# | column chart | data labels | show values | label position | OutsideEnd | Excel automation | chart series | .NET
// Common Searches: Aspose.Cells add data labels to column chart C# | display values on each bar in Aspose.Cells chart | set label position outside column chart Aspose.Cells | C# Aspose.Cells show series values in chart | how to enable value labels for column chart using Aspose.Cells
// Developer Intent: Attach a numeric label to every column in a chart generated with Aspose.Cells so the exact value is visible.
// Use Cases: Produce a sales‑by‑region report where each column bar is annotated with its revenue figure for instant visual comparison. | Build a performance dashboard that places precise metric values next to each column, improving readability on crowded charts. | Automate financial statements that require audit‑ready column charts with amount labels displayed on every bar.
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart and turns on value labels positioned outside each column. | Show how to customize font size, color, and background of data labels for a column chart series in Aspose.Cells. | Explain how to hide data labels for selected points while keeping them visible for the remaining columns in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills A1:B4 with categories and numbers, inserts a column chart, binds the series, enables data labels to display each point’s value, positions them outside the columns, and saves the file as XLSX.
    public class ColumnChartDataLabelsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(45);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series and show the exact values
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true; // Show value for each point
                series.DataLabels.Position = LabelPositionType.OutsideEnd; // Position labels outside the columns

                // Save the workbook with the chart
                workbook.Save("ColumnChartWithDataLabels.xlsx");
                Console.WriteLine("Workbook saved successfully as ColumnChartWithDataLabels.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ColumnChartDataLabelsDemo.Run();
        }
    }
}
