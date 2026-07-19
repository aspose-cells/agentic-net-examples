// Title: Aspose.Cells for .NET – Set Category Axis Tick Labels to Stacked (Vertical)
// Description: Demonstrates how to create a workbook, add a column chart, and use the ChartTextDirectionType.Stacked enumeration to display category‑axis tick labels in a vertical stacked layout before saving the file.
// Keywords: Aspose.Cells tick label orientation | ChartTextDirectionType.Stacked example | vertical stacked axis labels .NET | C# Aspose.Cells chart formatting | Excel chart label direction programmatically
// Common Searches: Aspose.Cells set category axis labels stacked | ChartTextDirectionType stacked C# | vertical tick labels Aspose.Cells chart | how to change axis label direction in Aspose.Cells | stacked tick labels Excel chart .NET
// Developer Intent: Apply the ChartTextDirectionType.Stacked enum to render category‑axis tick labels in a vertical stacked arrangement.
// Use Cases: Improve readability of narrow column charts by stacking axis labels vertically. | Generate Excel reports where overlapping category names must be avoided. | Create compact chart layouts for dashboards that require space‑efficient labeling.
// AI Prompts: Generate C# code with Aspose.Cells that sets the category axis tick labels to stacked for a line chart. | Explain the impact of ChartTextDirectionType values on tick label orientation and show how to toggle between Stacked and Rotated. | Provide a step‑by‑step tutorial to build a chart, apply stacked tick label direction, and save the workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart, and use the ChartTextDirectionType.Stacked enumeration to display category‑axis tick labels in a vertical stacked layout before saving the file.
    public class TickLabelsStackedDirectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set tick labels direction to Stacked (vertical orientation)
                chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Stacked;

                // Save the workbook to a file
                workbook.Save("TickLabelsStackedDirectionDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TickLabelsStackedDirectionDemo.Run();
        }
    }
}
