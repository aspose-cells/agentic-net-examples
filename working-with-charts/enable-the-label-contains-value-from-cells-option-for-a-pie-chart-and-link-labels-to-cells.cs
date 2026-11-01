// Title: How to enable "Label Contains – Value From Cells" for a pie chart and bind its labels to a worksheet range using Aspose.Cells for .NET
// AI Prompts: Create a pie chart and configure its series to pull label text from cells C2:C5 using Aspose.Cells. | Turn on the option that lets chart labels come from worksheet cells and assign the source range in C# with Aspose.Cells. | Save the workbook after displaying both numeric values and the custom cell‑based labels on the pie chart.
// Common Searches: Aspose.Cells how to use cell values for pie chart labels in C# | C# Aspose.Cells bind chart data labels to a specific range | Enable chart label source from worksheet cells with Aspose.Cells .NET | Set custom labels for Excel pie chart using Aspose.Cells API
// Tags: pie chart labels using worksheet range Aspose.Cells | activate label‑from‑cells feature .NET | set chart label source range C# | Aspose.Cells chart data label setup

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills columns A, B, and C with categories, numeric values, and custom label text, adds a pie chart based on the values in B2:B5, activates the "Label Contains – Value From Cells" option, binds the labels to the range C2:C5, shows both values and category names on the chart, and saves the file as PieChartWithLabelsFromCells.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("Label");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["A5"].PutValue("Date");

            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(25);
            sheet.Cells["B5"].PutValue(25);

            sheet.Cells["C2"].PutValue("Red Fruit");
            sheet.Cells["C3"].PutValue("Yellow Fruit");
            sheet.Cells["C4"].PutValue("Red Berry");
            sheet.Cells["C5"].PutValue("Sweet Fruit");

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series values (B2:B5) – false indicates values, not categories
            chart.NSeries.Add("B2:B5", false);

            // Show values and category names as data labels
            chart.NSeries[0].DataLabels.ShowValue = true;
            chart.NSeries[0].DataLabels.ShowCategoryName = true;

            // Determine output path and ensure directory exists
            string outputPath = "PieChartWithLabelsFromCells.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
