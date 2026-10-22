// Title: Enable text wrapping for pie chart data labels with Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a pie chart using Aspose.Cells and sets DataLabels.IsTextWrapped to true for the series. | Show how to add data labels to a pie chart and turn on text wrapping for those labels with the Aspose.Cells .NET API. | Provide a complete example that saves an Excel workbook containing a pie chart whose data label text automatically wraps.
// Common Searches: Aspose.Cells C# wrap text in pie chart data labels | Set DataLabels.IsTextWrapped property for Excel pie chart using Aspose.Cells | Enable multiline data labels on a pie chart with Aspose.Cells .NET | How to make pie chart labels wrap in generated Excel file using Aspose.Cells
// Tags: pie chart data label wrapping Aspose.Cells | DataLabels.IsTextWrapped C# | Aspose.Cells enable chart label wrap | Excel pie chart multiline labels .NET | Aspose.Cells chart formatting example

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a new workbook, fills cells with category and value data, adds a pie chart, configures its series and categories, shows values on the first series, enables DataLabels.IsTextWrapped to wrap label text, ensures the output directory exists, and saves the file as PieChartWithWrappedLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate data for the pie chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("Apple");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["A3"].PutValue("Banana");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("Cherry");
            ws.Cells["B4"].PutValue(50);

            // Add a pie chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Define the series and categories for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Show data labels on the first series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Enable text wrapping for data labels
            chart.NSeries[0].DataLabels.IsTextWrapped = true;

            // Define output file path
            string outputPath = "PieChartWithWrappedLabels.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
