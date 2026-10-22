// Title: How to right‑align data label text in a horizontal bar chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that sets the data label alignment to right for a horizontal bar chart series. | Update an existing Aspose.Cells chart example to right‑justify the data labels of a bar chart and save the workbook.
// Common Searches: Aspose.Cells C# set data label alignment right in bar chart | right justify chart data labels in a horizontal bar chart using Aspose.Cells .NET | how to change data label text alignment for a bar chart series in Aspose.Cells | C# Aspose.Cells align data labels to the right in a stacked bar chart
// Tags: Aspose.Cells set data label alignment C# | horizontal bar chart label formatting Aspose.Cells | chart series data label justification .NET | C# Aspose.Cells bar chart label position | customize chart data labels Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills cells A1:B4 with categories and values, adds a horizontal bar chart, binds the series to the data, enables data labels, sets the label alignment to right, and saves the file as HorizontalBarChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the horizontal bar chart
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["A3"].PutValue("B");
            cells["A4"].PutValue("C");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);

            // Add a horizontal bar chart (Bar chart type)
            // Parameters: chart type, upper‑left row, upper‑left column, lower‑right row, lower‑right column
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 15, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].XValues = "A2:A4";

            // Enable data labels for the series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Save the workbook to a file
            workbook.Save("HorizontalBarChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
