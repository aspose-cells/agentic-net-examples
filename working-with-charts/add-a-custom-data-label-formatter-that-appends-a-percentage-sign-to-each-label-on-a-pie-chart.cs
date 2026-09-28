// Title: Add a percentage sign to pie chart data labels with a custom number format using Aspose.Cells for .NET
// AI Prompts: Create a pie chart in Aspose.Cells and configure its data labels to display each value followed by a % sign. | Apply a custom number format to chart data labels in C# so the labels appear as percentages. | Generate an Excel workbook with a pie chart where the data label formatter appends a percent symbol to every label.
// Common Searches: how to show percent sign on pie chart data labels in Aspose.Cells C# | Aspose.Cells custom number format for chart labels .NET | C# example adding % to pie chart labels using Aspose.Cells | format pie chart data labels as percentages in Excel with Aspose.Cells | Aspose.Cells chart data label formatter percentage sign
// Tags: pie chart data label number format Aspose.Cells | custom percentage label formatter .NET | Aspose.Cells chart data label formatting | Excel pie chart label percent sign C# | Aspose.Cells set data label format

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a new workbook, fills it with category and value data, adds a pie chart, links the series and category ranges, enables data labels, applies the number format "0%" to append a percentage sign to each label, and saves the file as PieChartWithCustomLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Get the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the pie chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(70);

            // Add a pie chart to the worksheet
            int chartIdx = ws.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set the series data range (values)
            chart.NSeries.Add("B2:B3", true);

            // Set the category (label) data range for the series
            // Use XValues for compatibility across Aspose.Cells versions
            chart.NSeries[0].XValues = "A2:A3";

            // Enable data labels to show the values
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Apply a custom number format that appends a percentage sign
            chart.NSeries[0].DataLabels.NumberFormat = "0%";

            // Save the workbook with the chart
            string outputPath = "PieChartWithCustomLabels.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
