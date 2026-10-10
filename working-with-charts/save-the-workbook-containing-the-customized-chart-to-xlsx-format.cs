// Title: Save a workbook with a customized column chart to XLSX using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, adds sample data, inserts a column chart with a right‑aligned legend, and saves the file as an XLSX document using Aspose.Cells. | Show how to set the chart title, define the series range, and ensure the output folder exists before calling Workbook.Save with SaveFormat.Xlsx in C#. | Write a self‑contained Aspose.Cells example that builds a worksheet, populates rows, adds a column chart, customizes its legend placement, and writes the workbook to a specified .xlsx path.
// Common Searches: Aspose.Cells C# create column chart and export workbook to XLSX file | how to set legend position on a chart using Aspose.Cells .NET | save workbook to specific folder with Aspose.Cells ensuring directory exists | example of adding data and chart to worksheet then saving as .xlsx in C#
// Tags: Aspose.Cells add column chart | Aspose.Cells customize chart legend placement | Aspose.Cells export workbook to XLSX | Aspose.Cells ensure output directory exists | Aspose.Cells populate worksheet data for chart

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The program creates a new workbook, fills it with sample data, adds a column chart whose legend is positioned on the right, guarantees the target directory exists, and saves the workbook as CustomizedChart.xlsx in XLSX format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sample Column Chart";

            // Set the data range for the series (values)
            chart.NSeries.Add("B2:B4", true);
            // Category data can be set if the API supports it; omitted here for compatibility
            // chart.NSeries[0].CategoryData = "A2:A4";

            // Example customization: place the legend on the right
            chart.Legend.Position = LegendPositionType.Right;

            // Determine output path and ensure directory exists
            string outputPath = "CustomizedChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook in XLSX format
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
