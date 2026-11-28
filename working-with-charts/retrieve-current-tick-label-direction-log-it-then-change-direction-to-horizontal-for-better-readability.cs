// Title: Read current tick label direction and set it to horizontal on a chart’s category axis with Aspose.Cells for .NET (C#)
// Common Searches: how to get the current tick label direction of a chart in Aspose.Cells C# | changing chart category axis tick labels to horizontal using Aspose.Cells for .NET | Aspose.Cells C# example to read and modify tick label orientation | retrieve and set ChartTextDirectionType for axis tick labels in a .xlsx file | C# code to log and update chart axis tick label direction with Aspose.Cells
// Tags: chart category axis tick label direction Aspose.Cells | set tick label orientation horizontal Aspose.Cells | retrieve chart tick label DirectionType C# | modify axis text direction .NET Aspose.Cells | Aspose.Cells chart tick labels DirectionType example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace TickLabelDirectionDemo
{
    // The sample creates a workbook, adds a column chart, accesses the category axis tick labels, logs the existing DirectionType, changes it to Horizontal for better readability, verifies the update, and saves the file as TickLabelDirectionDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(150);
                worksheet.Cells["B4"].PutValue(180);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the category axis tick labels
                Axis categoryAxis = chart.CategoryAxis;
                TickLabels tickLabels = categoryAxis.TickLabels;

                // Retrieve and log the current direction
                ChartTextDirectionType currentDirection = tickLabels.DirectionType;
                Console.WriteLine($"Current Tick Labels Direction: {currentDirection}");

                // Change the direction to Horizontal for better readability
                tickLabels.DirectionType = ChartTextDirectionType.Horizontal;

                // Verify the change
                Console.WriteLine($"New Tick Labels Direction: {tickLabels.DirectionType}");

                // Save the workbook
                string outputPath = "TickLabelDirectionDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
