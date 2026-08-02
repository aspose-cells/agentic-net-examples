// Title: Remove a Chart Shape from a Chart Sheet with ShapeCollection.RemoveAt in Aspose.Cells for .NET
// Description: Creates a workbook, adds a chart sheet, inserts a column chart, accesses the sheet's Shapes collection, removes the chart shape using ShapeCollection.RemoveAt, and saves the file as ChartSheetWithoutChart.xlsx.
// Keywords: Aspose.Cells remove chart shape | ShapeCollection.RemoveAt .NET | delete chart from worksheet Aspose.Cells | chart sheet shape removal C# | Aspose.Cells chart control deletion
// Common Searches: how to delete a chart shape using Aspose.Cells | remove chart from chart sheet ShapeCollection.RemoveAt | Aspose.Cells .NET delete chart programmatically | remove first chart shape in worksheet C# | Aspose.Cells remove chart control example
// Developer Intent: Remove the chart control from a chart sheet by deleting its shape from the worksheet's Shapes collection.
// Use Cases: Clean up temporary charts before exporting a financial report workbook. | Generate a template workbook that contains only raw data by stripping out chart objects. | Convert a chart sheet to a data‑only sheet for downstream processing or analysis.
// AI Prompts: Show how to confirm that a chart shape was removed after calling ShapeCollection.RemoveAt in Aspose.Cells. | Provide code to delete a chart by its name instead of using an index with Aspose.Cells for .NET. | Explain the impact of ShapeCollection.RemoveAt on the underlying Chart object and overall workbook size.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a chart sheet, inserts a column chart, accesses the sheet's Shapes collection, removes the chart shape using ShapeCollection.RemoveAt, and saves the file as ChartSheetWithoutChart.xlsx.
class RemoveChartFromChartSheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Add a regular worksheet (will host a chart) and get the worksheet reference.
            Worksheet chartSheet = workbook.Worksheets.Add("MyChartSheet");

            // Add a chart to the worksheet.
            // Parameters: upper left row, upper left column, lower right row, lower right column, chart type.
            // Cast ChartType to int for compatibility with older Aspose.Cells versions.
            int chartIndex = chartSheet.Charts.Add(0, 0, 10, 5, (int)ChartType.Column);
            Chart chart = chartSheet.Charts[chartIndex];

            // The chart is also represented as a shape in the sheet's Shapes collection.
            ShapeCollection shapes = chartSheet.Shapes;

            // If there is at least one shape (the chart), remove it.
            if (shapes.Count > 0)
            {
                shapes.RemoveAt(0); // Remove the chart control.
            }

            // Save the workbook to verify the chart has been removed.
            string outputPath = "ChartSheetWithoutChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
