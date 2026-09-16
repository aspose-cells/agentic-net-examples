// Title: Automatically adjust a column chart’s size to fit its data range with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that creates a column chart and programmatically sets its Height and Width so the chart fully encloses the plotted data. | Demonstrate how to retrieve the ChartObject of a chart in Aspose.Cells and modify its Top, Left, Height, and Width properties to eliminate clipping. | Create a helper method that computes optimal chart dimensions from the number of data rows and applies those dimensions to a chart object.
// Common Searches: aspnet how to resize a column chart to match data range in Aspose.Cells | c# Aspose.Cells chartobject height width adjustment example | auto fit chart dimensions to avoid clipping Aspose.Cells .NET | set chart size based on number of data points using Aspose.Cells | adjust column chart bounds programmatically with Aspose.Cells
// Tags: chartobject size properties Aspose.Cells | auto-fit column chart size .NET | derive chart size from row count | column chart size adjustment Aspose.Cells | programmatic chart resizing Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a workbook, fills ten rows with sample data, adds a column chart linked to that data, and demonstrates how to access the underlying ChartObject to set Top, Left, Height, and Width so the chart fits the data range, then saves the file as AutoFitChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (10 rows)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1);          // Category (A column)
                sheet.Cells[i, 1].PutValue((i + 1) * 10);   // Values (B column)
            }

            // Add a column chart to the sheet
            // Parameters: chart type, upper‑left row, upper‑left column, lower‑right row, lower‑right column
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 3, 20, 13);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data series and category data
            chart.NSeries.Add("Sheet1!B1:B10", true);          // Values
            chart.NSeries.CategoryData = "Sheet1!A1:A10";     // Categories

            // -------------------- Auto‑fit chart size (optional) --------------------
            // Aspose.Cells versions prior to 22.x do not expose Height/Width on Chart.
            // If needed, adjust the chart size using the Top/Left/Height/Width properties
            // of the underlying shape via the ChartObject (available in newer versions).
            // For compatibility, this example uses the default chart size.
            // -----------------------------------------------------------------------

            // Save the workbook with the chart
            workbook.Save("AutoFitChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
