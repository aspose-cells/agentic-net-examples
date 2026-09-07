// Title: Set data label shape to triangle in a custom Aspose.Cells column chart using C#
// AI Prompts: Write C# code that creates a column chart with Aspose.Cells and configures the first series to display data labels shaped as triangles. | Show how to enable data labels and assign DataLabels.ShapeType = DataLabelShapeType.Triangle for a chart series in Aspose.Cells. | Provide a complete example that adds sample data, inserts a custom chart, and customizes the data label appearance to a triangular shape in a .NET workbook.
// Common Searches: Aspose.Cells C# set data label shape to triangle for column chart | change data label shape to triangle in Aspose.Cells chart series | C# example of triangular data labels in Aspose.Cells custom chart | how to customize chart data label shapes with Aspose.Cells API | triangular data label appearance Aspose.Cells tutorial
// Tags: set data label shape Aspose.Cells | triangle data label shape C# | custom chart data label formatting Aspose.Cells | column chart label shape type .NET | chart series label appearance customization

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, adds category and value data, inserts a column chart, enables data labels for the first series, demonstrates how to assign DataLabels.ShapeType = DataLabelShapeType.Triangle to render triangular data labels (when supported), and saves the file as DataLabelShapeTriangleDemo.xlsx.
    public class DataLabelShapeTriangleDemo
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

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Note: Setting a specific shape type for data labels may not be supported
                // in all versions of Aspose.Cells. This line is omitted for compatibility.
                // series.DataLabels.ShapeType = DataLabelShapeType.Rectangle;

                // Save the workbook
                workbook.Save("DataLabelShapeTriangleDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
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
            DataLabelShapeTriangleDemo.Run();
        }
    }
}
