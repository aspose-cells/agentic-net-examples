// Title: Aspose.Cells C# Example: Attempt to Set Triangle Shape for Chart Data Labels (Unsupported)
// Description: This C# sample creates a workbook, adds a column chart with sample data, enables data labels to show values, and explains that Aspose.Cells for .NET does not provide an API to change the data‑label shape to a triangle, so the default shape is retained before saving the file.
// Keywords: Aspose.Cells data label shape | C# chart data label triangle | Aspose.Cells custom chart example | Aspose.Cells label shape limitation | Aspose.Cells column chart C#
// Common Searches: Aspose.Cells change data label shape | C# set triangle data label Aspose.Cells | chart data label customization Aspose.Cells | unsupported chart label shape Aspose.Cells | how to modify data label appearance in Aspose.Cells
// Developer Intent: The developer wants to apply a triangle shape to the data labels of a chart series using Aspose.Cells for .NET.
// Use Cases: Generate a column chart with category and value series and display label values. | Demonstrate the current limitation of Aspose.Cells regarding data‑label shape customization. | Save the resulting workbook with the chart to an Excel file.
// AI Prompts: Write C# code with Aspose.Cells that builds a column chart, shows data label values, and includes a comment about the unsupported triangle shape. | Explain why Aspose.Cells cannot change data label shapes and suggest possible work‑arounds or alternative libraries. | Create a step‑by‑step tutorial for adding data labels to a chart in Aspose.Cells and clarify the limitation on custom label shapes.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# sample creates a workbook, adds a column chart with sample data, enables data labels to show values, and explains that Aspose.Cells for .NET does not provide an API to change the data‑label shape to a triangle, so the default shape is retained before saving the file.
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

                // Note: Aspose.Cells does not provide a direct API to change the shape of data labels.
                // The original code attempted to set a triangle shape, which is not supported.
                // Therefore, this step is omitted.

                // Save the workbook
                workbook.Save("DataLabelShapeTriangleDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
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
