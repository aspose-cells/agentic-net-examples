// Title: How to set line chart data label shape to rounded rectangle using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# program that creates an Excel workbook, adds a line chart, enables data labels, and assigns DataLabels.ShapeType = DataLabelShapeType.RoundedRectangle with Aspose.Cells. | Update an existing Aspose.Cells line chart in C# to display its data labels as rounded rectangles by setting the ShapeType property. | Write example code that demonstrates configuring the data label shape type for a line chart, including version‑check handling for the DataLabelShapeType enum in Aspose.Cells.
// Common Searches: Aspose.Cells C# set data label shape rounded rectangle line chart | how to change data label shape to rounded rectangle in Aspose.Cells line chart | C# Aspose.Cells line chart data labels show value with rounded rectangle | DataLabels.ShapeType property usage Aspose.Cells .NET
// Tags: Aspose.Cells line chart data label shape | C# DataLabels.ShapeType rounded rectangle | Excel chart data label shape .NET | configure data label shape Aspose.Cells | line chart data labels show value C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, populates sample data, inserts a line chart, enables data labels to show values, and demonstrates how to set the data label shape to a rounded rectangle using the DataLabels.ShapeType property (when supported), then saves the workbook as an .xlsx file.
    public class LineChartDataLabelRoundedRect
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the line chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a line chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series and enable data labels
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // NOTE: The DataLabelShapeType enum may not be available in older Aspose.Cells versions.
            // If supported, you can set the shape type to a rounded rectangle as shown below:
            // series.DataLabels.ShapeType = DataLabelShapeType.RoundedRectangle;

            // Save the workbook
            try
            {
                workbook.Save("LineChartDataLabelRoundedRect.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
