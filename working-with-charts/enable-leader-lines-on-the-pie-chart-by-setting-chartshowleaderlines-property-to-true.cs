// Title: Enable leader lines on a pie chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a pie chart with Aspose.Cells and turns on leader lines by setting Chart.ShowLeaderLines to true. | Generate a minimal Aspose.Cells example that adds a pie chart, enables leader lines, and saves the workbook as an .xlsx file. | Show how to modify an existing Aspose.Cells pie chart to display leader lines in a .NET application.
// Common Searches: Aspose.Cells C# enable leader lines for Excel pie chart | How to turn on leader lines in a pie chart with Aspose.Cells .NET | Set ShowLeaderLines property on Aspose.Cells pie chart example | C# Aspose.Cells add leader lines to pie chart for Excel file | Display leader lines in pie chart using Aspose.Cells library
// Tags: pie chart leader lines Aspose.Cells | Chart.ShowLeaderLines C# Aspose.Cells | Aspose.Cells enable leader lines Excel chart | add leader lines to pie chart .NET | Aspose.Cells chart formatting leader lines | Excel pie chart leader lines Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new Workbook, fills cells A1:B5 with fruit categories and values, adds a pie chart, assigns the value range B2:B5 to the series, sets Chart.ShowLeaderLines = true to display leader lines, adds a title, and saves the file as PieChart_With_LeaderLines.xlsx.
class EnableLeaderLinesInPieChart
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
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(25);
            sheet.Cells["A5"].PutValue("Date");
            sheet.Cells["B5"].PutValue(25);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (values)
            chart.NSeries.Add("B2:B5", true);
            // Category data is taken from the first column automatically for pie charts

            // Optional: set a title for clarity
            chart.Title.Text = "Fruit Distribution";

            // Determine output file path
            string outputPath = "PieChart_With_LeaderLines.xlsx";

            // Ensure the output directory exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
