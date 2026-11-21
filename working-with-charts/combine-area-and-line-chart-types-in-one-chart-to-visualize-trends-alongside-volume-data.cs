// Title: Generate a combined Area and Line chart in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: how to add a line series to an area chart with Aspose.Cells C# | Aspose.Cells create mixed chart types area and line example | C# code to set date column as category axis in Aspose.Cells chart | override chart series type to line in Aspose.Cells workbook | save Excel workbook with combined area and line chart using Aspose.Cells
// Tags: area chart with line series Aspose.Cells | set category axis using date column Aspose.Cells | override series to line type Aspose.Cells | export Excel workbook with mixed chart Aspose.Cells | create combined chart programmatically C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills it with dates, volume, and price data, adds an Area chart, overrides the second series to a Line type, assigns the Date column as the category axis, sets a chart title, and saves the file as CombinedChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Date, Volume (area), Price (line)
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Volume");
            sheet.Cells["C1"].PutValue("Price");

            DateTime startDate = new DateTime(2023, 1, 1);
            Random rnd = new Random();
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i)); // Date
                sheet.Cells[i + 1, 1].PutValue(rnd.Next(100, 500));   // Volume
                sheet.Cells[i + 1, 2].PutValue(rnd.Next(20, 80));     // Price
            }

            // Add a combined chart (area + line) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Area, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // First series: Volume displayed as Area
            chart.NSeries.Add("B2:B11", true);
            chart.NSeries[0].Name = "Volume";

            // Second series: Price displayed as Line
            chart.NSeries.Add("C2:C11", true);
            chart.NSeries[1].Name = "Price";
            chart.NSeries[1].Type = ChartType.Line; // Override series type to Line

            // Set category (X) axis to the Date column
            chart.NSeries.CategoryData = "A2:A11";

            // Optional: set chart title
            chart.Title.Text = "Volume (Area) and Price (Line)";

            // Determine output file path and ensure directory exists
            string outputPath = "CombinedChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
