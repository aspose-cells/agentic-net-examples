// Title: Saving a modified workbook with updated chart to a new Excel file (ChartReport.xlsx) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook after updating chart series to a new file | Create placeholder chart when none exists and export workbook as ChartReport.xlsx with Aspose.Cells | Ensure output directory exists before calling Workbook.Save in Aspose.Cells .NET | Load Input.xlsx if present, modify chart data range, and save as a different Excel file using Aspose.Cells | Update chart data range in Aspose.Cells and write the result to a new workbook
// Tags: Workbook.Save to new Excel file Aspose.Cells | update chart series values Aspose.Cells C# | add placeholder chart when worksheet has no charts Aspose.Cells | create output directory before saving Aspose.Cells | conditional workbook load Input.xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads Input.xlsx if it exists; otherwise it creates a new workbook, adds a placeholder column chart, writes numeric values and category labels to cells A1:A4 and B1:B4, updates the first chart series to reference the new values, ensures the output directory exists, and saves the modified workbook as ChartReport.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "ChartReport.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";

                // Add a placeholder chart so subsequent code has a chart to modify.
                int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart placeholder = ws.Charts[chartIdx];
                placeholder.NSeries.Add("A1:A4", true);
            }

            // ----- Begin chart updates -----
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart.
            if (sheet.Charts.Count == 0)
            {
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart newChart = sheet.Charts[chartIdx];
                newChart.NSeries.Add("A1:A4", true);
            }

            Chart chart = sheet.Charts[0];

            // Write new data to worksheet cells.
            double[] values = { 10, 20, 30, 40 };
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };

            for (int i = 0; i < values.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(values[i]);      // Column A for values
                sheet.Cells[i, 1].PutValue(categories[i]); // Column B for categories
            }

            // Update the series to use the new ranges.
            chart.NSeries[0].Values = "A1:A4";
            // CategoryData property is not available in this version; categories will be shown as default labels.

            // ----- End chart updates -----

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
