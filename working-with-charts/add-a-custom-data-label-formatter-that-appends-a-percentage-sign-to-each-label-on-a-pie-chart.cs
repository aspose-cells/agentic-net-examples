// Title: Add percentage sign to pie chart data labels using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a pie chart, turns on data labels, and formats them to show values as percentages. | Demonstrate how to assign the number format "0%" to chart series data labels in an Aspose.Cells workbook using .NET. | Provide a full example that builds sample data, adds a pie chart, enables label display, and applies a percentage format to each label in C#.
// Common Searches: aspocells c# format pie chart labels as percent | set data label number format to 0% in Aspose.Cells chart | display pie slice values with % using Aspose.Cells .NET | how to show percentages on Excel pie chart labels with Aspose.Cells
// Tags: Aspose.Cells pie chart label formatting | C# Excel chart label number format | Aspose.Cells data label customization | percentage display on pie chart labels | Aspose.Cells workbook chart settings

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a pie chart, enables data labels, applies the "0%" number format so each slice shows its value with a trailing percent sign, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart series (values)
            chart.NSeries.Add("B2:B4", true);

            // (Optional) Set category labels for the pie slices.
            // For pie charts, category data can be omitted; Excel will use the first column by default.
            // If needed, you can uncomment the line below and ensure your Aspose.Cells version supports it.
            // chart.NSeries[0].CategoryData = "A2:A4";

            // Show data labels on the pie slices
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Apply a custom number format that appends a percentage sign
            chart.NSeries[0].DataLabels.NumberFormat = "0%";

            // Save the workbook to a file
            string outputPath = "PieChartWithCustomLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
