// Title: Add a column chart in Aspose.Cells for .NET and color its series with the workbook’s Accent5 theme color
// AI Prompts: Generate C# code using Aspose.Cells to create a column chart and set each series’ fill and marker colors to the workbook’s Accent5 theme color. | Show how to retrieve a workbook’s Accent5 theme color with GetThemeColor and apply it to chart series in Aspose.Cells. | Provide a complete example that builds sample data, adds a column chart, applies the Accent5 theme color to series, and saves the workbook as an .xlsx file.
// Common Searches: aspnet aspocells how to set chart series color to workbook accent5 theme | c# Aspose.Cells get theme color Accent5 for chart series | apply workbook theme colors to column chart series using Aspose.Cells .NET | set marker and fill color of chart series from workbook theme in Aspose.Cells | example of using GetThemeColor with chart series in Aspose.Cells
// Tags: Aspose.Cells set chart series theme color | C# GetThemeColor Accent5 | column chart series fill color Aspose.Cells | apply workbook theme to chart Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, populates sample data, inserts a column chart, retrieves the workbook's Accent5 theme color via GetThemeColor, applies that color to each series' fill and marker, and saves the file as an .xlsx workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value1");
            sheet.Cells["C1"].PutValue("Value2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];

            // Add series for the two value columns
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.Add("C2:C4", true);

            // Retrieve the workbook's Accent5 theme color
            Color accent5 = workbook.GetThemeColor(ThemeColorType.Accent5);

            // Apply Accent5 color to each series
            foreach (Series series in chart.NSeries)
            {
                // Set the fill color of the series
                series.Area.ForegroundColor = accent5;
                // Set the marker color (if markers are used)
                series.Marker.ForegroundColor = accent5;
            }

            // Define output file path
            string outputPath = "ChartWithAccent5.xlsx";

            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (string.IsNullOrEmpty(directory))
            {
                directory = Directory.GetCurrentDirectory();
            }
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
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
