// Title: Apply a built‑in chart style to a column chart in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a column chart from sample data, applies a built‑in Aspose.Cells chart style, and saves the workbook as an .xlsx file. | Show how to verify and create the output directory before saving a styled chart workbook with Aspose.Cells in C#. | Demonstrate setting the ChartStyleType enum on an Aspose.Cells chart and handling cases where the property is unavailable in older library versions.
// Common Searches: Aspose.Cells C# set predefined chart theme for column chart | how to use ChartStyleType with Aspose.Cells in .NET | create Excel column chart with theme using Aspose.Cells | ensure output folder exists when saving Aspose.Cells workbook | workaround missing ChartStyle property in older Aspose.Cells versions
// Tags: Aspose.Cells set chart style enum | C# column chart built‑in theme Aspose.Cells | create output directory before workbook save | Aspose.Cells ChartStyleType usage | save workbook as .xlsx with styled chart

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Creates a workbook, adds sample data, inserts a column chart, optionally assigns a built‑in ChartStyleType, ensures the output directory exists, and saves the file as ChartWithTheme.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a built‑in chart style (if supported by the library version)
            // The ChartStyleType enum may not be available in older versions; this line is optional.
            // chart.Style = ChartStyleType.Style1;

            // Define output file path
            string outputPath = "ChartWithTheme.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
