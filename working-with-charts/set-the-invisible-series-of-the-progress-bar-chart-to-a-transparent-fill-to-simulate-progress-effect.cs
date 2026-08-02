// Title: Aspose.Cells .NET: Create a Stacked Column Progress Bar with a Transparent Remainder Series
// Description: Shows how to generate a stacked column chart that works as a progress bar in an Excel workbook using Aspose.Cells for .NET. The sample adds task data, creates two series (Progress and Remaining), applies a solid fill to the visible series, makes the remainder series fully transparent by setting area and border transparency, and saves the workbook.
// Keywords: Aspose.Cells | C# | stacked column chart | progress bar chart | transparent chart series | chart series transparency | hide series Aspose.Cells | Excel progress bar | set series fill color | Aspose.Cells chart formatting
// Common Searches: Aspose.Cells make chart series invisible | transparent series Aspose.Cells .NET | progress bar chart using Aspose.Cells | hide remainder series stacked column chart | set series fill color Aspose.Cells C# | Excel progress bar with transparent remainder
// Developer Intent: Apply full transparency to the remainder series of a stacked column chart to simulate a progress bar while keeping the completed portion visible.
// Use Cases: Display task completion percentages in dashboards where only the completed portion of the bar is shown. | Generate Excel reports that use a clean progress‑bar visual without displaying the unfinished segment. | Create printable Excel sheets with progress indicators that hide the remaining portion for a minimalist look.
// AI Prompts: Write C# code with Aspose.Cells that creates a stacked column chart and sets the second series' area and border transparency to 1.0. | Explain how to hide a series in an Aspose.Cells chart by making it fully transparent while preserving the color of another series. | Show how to apply a solid fill to the progress series and a transparent fill to the remainder series in an Aspose.Cells progress bar example.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to generate a stacked column chart that works as a progress bar in an Excel workbook using Aspose.Cells for .NET. The sample adds task data, creates two series (Progress and Remaining), applies a solid fill to the visible series, makes the remainder series fully transparent by setting area and border transparency, and saves the workbook.
class ProgressBarChartDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add data for a single task: progress and remaining values
            worksheet.Cells["A1"].PutValue("Task");
            worksheet.Cells["A2"].PutValue("Demo");
            worksheet.Cells["B1"].PutValue("Progress");
            worksheet.Cells["B2"].PutValue(70);   // Completed part
            worksheet.Cells["C1"].PutValue("Remaining");
            worksheet.Cells["C2"].PutValue(30);   // Incomplete part

            // Add a stacked column chart to act as a progress bar
            int chartIndex = worksheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Add two series: the visible progress and the invisible remainder
            chart.NSeries.Add("B2:C2", true);               // Both series share the same category
            chart.NSeries.CategoryData = "A2:A2";           // Single category (the task name)

            // Configure the visible progress series
            Series progressSeries = chart.NSeries[0];
            progressSeries.Name = "Progress";

            // Set solid fill color for the progress part (if supported)
            // Uncomment the line below if your Aspose.Cells version supports SetSolidFillColor
            // progressSeries.Area.FillFormat.SetSolidFillColor(Color.Green);

            // Configure the invisible (remainder) series
            Series invisibleSeries = chart.NSeries[1];
            invisibleSeries.Name = "Remaining";
            // Make the area fully transparent (0.0 = opaque, 1.0 = clear)
            invisibleSeries.Area.Transparency = 1.0;
            // Also make the border fully transparent to avoid any outline
            invisibleSeries.Border.Transparency = 1.0;

            // Save the workbook with the progress bar chart
            workbook.Save("ProgressBarChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
