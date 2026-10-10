// Title: How to set vertical axis tick labels to stacked orientation using ChartTextDirectionType in Aspose.Cells C#
// AI Prompts: Add a statement that assigns ChartTextDirectionType.Stacked to chart.ValueAxis.TickLabels.TextDirection so the vertical axis labels are displayed stacked. | Update the chart configuration to change the tick‑label direction of the value axis to stacked by using the ChartTextDirectionType enumeration in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells set value axis tick label direction stacked C# | ChartTextDirectionType.Stacked example Aspose.Cells .NET | vertical axis labels stacked orientation Aspose.Cells column chart | how to change tick label orientation in Aspose.Cells chart | C# Aspose.Cells chart label direction stacked
// Tags: ChartTextDirectionType stacked tick labels | Aspose.Cells set value axis label direction | C# column chart vertical label orientation | Aspose.Cells chart axis text direction | Excel chart stacked tick labels Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds sample data, inserts a column chart, sets the value axis title, configures the value axis tick labels to use the stacked orientation via ChartTextDirectionType, and saves the file as 'ChartWithStackedTickLabels.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the chart data range
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set the value axis title (use Title.Text property)
                chart.ValueAxis.Title.Text = "Value Axis";

                // Determine output file path
                string outputFile = "ChartWithStackedTickLabels.xlsx";

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputFile)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
