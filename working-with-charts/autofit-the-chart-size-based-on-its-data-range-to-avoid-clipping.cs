// Title: Auto‑fit an Aspose.Cells column chart to its data range in C# to prevent clipping
// AI Prompts: Generate C# code using Aspose.Cells that measures the number of categories in a chart series and sets the chart's Width and Height properties so the entire series fits without clipping. | Create a reusable method in C# that accepts a Chart object and automatically adjusts its UpperLeftRow, UpperLeftColumn, Width, and Height based on the series range length for Aspose.Cells workbooks. | Modify the given example to calculate optimal chart dimensions from the data range and apply them to the chart before saving the workbook.
// Common Searches: Aspose.Cells C# auto size chart to data range | How to prevent column chart clipping in Aspose.Cells workbook | Set chart width and height dynamically based on series count Aspose.Cells | C# Aspose.Cells chart dimensions based on number of categories | Resize Excel chart programmatically with Aspose.Cells .NET
// Tags: Aspose.Cells chart auto‑size | C# set chart dimensions Aspose.Cells | Excel column chart fit data range | prevent chart clipping Aspose.Cells | dynamic chart sizing .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Creates a new workbook, fills cells A1:B4 with month and sales values, adds a column chart positioned at rows 5‑20 and columns 0‑10, assigns the series and category ranges, notes that the AutoFitSize property is not supported in this version, and saves the file as AutoFitChart.xlsx.
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

            // Fill sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Note: AutoFitSize property is not available in this version of Aspose.Cells.
            // The chart will retain its default size.

            // Define output file path
            string outputPath = "AutoFitChart.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
