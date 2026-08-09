// Title: Aspose.Cells .NET – Triangle Shape for Chart Data Labels (Unsupported)
// Description: This example creates a workbook, adds sample data, inserts a column chart, enables data labels, and explains that Aspose.Cells does not support setting a triangle shape for chart data labels before saving the file.
// Keywords: Aspose.Cells | C# | .NET | chart data labels | triangle shape | unsupported feature | column chart | workbook example | chart customization | Aspose.Cells API
// Common Searches: Aspose.Cells set data label shape | triangle data label Aspose.Cells .NET | custom chart data label shape support | how to change chart data label shape in Aspose.Cells | Aspose.Cells chart label shape limitation
// Developer Intent: Set the data label shape to a triangle for a custom chart using Aspose.Cells.
// Use Cases: Generate a column chart with visible data labels showing series values. | Attempt to apply a triangle shape to data labels and handle the lack of API support. | Save the workbook after configuring chart and data label settings.
// AI Prompts: Provide C# code that creates a column chart with data labels using Aspose.Cells and explains why a triangle shape cannot be applied. | Suggest alternative visual cues (e.g., marker styles or custom shapes) to emphasize data points when triangle data labels are unavailable in Aspose.Cells. | Write a try‑catch block that logs a clear message when an unsupported data label shape is requested in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds sample data, inserts a column chart, enables data labels, and explains that Aspose.Cells does not support setting a triangle shape for chart data labels before saving the file.
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

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Note: DataLabel shape types (e.g., triangle) are not supported in Aspose.Cells.
                // The following line has been removed to avoid compilation errors.

                // Save the workbook
                workbook.Save("DataLabelShapeTriangleDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
