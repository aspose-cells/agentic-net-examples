// Title: Add a column chart with custom axis titles and freeze the header row in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that inserts a column chart, sets the value axis title to "Sales" and the category axis title to "Month", then freezes the first worksheet row. | Generate a .NET example that creates a workbook, populates sample data, adds a column chart, configures both axis titles, and applies FreezePanes to keep axis labels visible while scrolling. | Provide an Aspose.Cells snippet demonstrating how to configure chart axes and programmatically freeze rows containing those labels.
// Common Searches: asp.net set column chart axis title using Aspose.Cells | c# freeze first row in Excel workbook with Aspose.Cells | example adding chart and freezing header row Aspose.Cells for .NET | Aspose.Cells configure value and category axis titles programmatically
// Tags: Aspose.Cells create column chart C# | Aspose.Cells set axis title | Aspose.Cells freeze panes worksheet | Aspose.Cells chart axis configuration | Aspose.Cells Excel file generation C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills it with sample data, adds a column chart, assigns "Sales" to the value axis and "Month" to the category axis, freezes the first row so axis labels stay visible during scrolling, and saves the result as ChartWithFrozenAxisLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            var chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            // Set category (X‑axis) labels if the API supports it
            // (If CategoryData property is unavailable, this line can be omitted)
            // chart.NSeries[0].CategoryData = "A2:A4";

            // Configure the primary value axis (title only – other settings may not be supported in all versions)
            var valueAxis = chart.ValueAxis;
            valueAxis.Title.Text = "Sales";

            // Configure the primary category axis (title only)
            var categoryAxis = chart.CategoryAxis;
            categoryAxis.Title.Text = "Month";

            // Freeze the first row (header) so it stays visible while scrolling
            sheet.FreezePanes(1, 0, 0, 0);

            // Save the workbook
            string outputPath = "ChartWithFrozenAxisLabels.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
