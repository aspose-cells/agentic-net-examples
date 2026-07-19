// Title: C# – Add a Linear Trendline to an Aspose.Cells Chart and Get Its Name
// Description: Loads an existing workbook or creates a new one with sample X‑Y data, ensures a line chart exists, adds a linear trendline to the first series, assigns a name, saves the file, and prints the trendline identifier to the console using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart trendline | add linear trendline Aspose.Cells | retrieve trendline name .NET | create line chart programmatically | save workbook with trendline | chart series trendline API | trendline equation workaround | Aspose.Cells chart automation
// Common Searches: how to add a linear trendline to a chart with Aspose.Cells | get trendline name from Aspose.Cells chart series | create line chart if none exists Aspose.Cells C# | save workbook after modifying chart objects | Aspose.Cells trendline equation not exposed
// Developer Intent: Add a linear trendline to the first series of a chart and output its identifier.
// Use Cases: Automatically annotate generated line charts with a linear trendline for analytical reports. | Update existing workbooks by inserting a trendline, saving the changes, and logging the trendline name for downstream processing. | Create a new workbook with sample data, ensure a chart is present, apply a trendline, and persist the result for export.
// AI Prompts: Generate C# code that loads or creates a workbook, adds a line chart if missing, inserts a linear trendline on the first series, and prints the trendline name. | Explain why Aspose.Cells does not expose the trendline equation directly and suggest a practical workaround. | Provide best‑practice error handling for saving a workbook after chart modifications with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing workbook or creates a new one with sample X‑Y data, ensures a line chart exists, adds a linear trendline to the first series, assigns a name, saves the file, and prints the trendline identifier to the console using Aspose.Cells for .NET.
class TrendlineFormulaExtractor
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            Workbook workbook;
            Worksheet sheet;

            // Load existing workbook if it exists; otherwise create a new one with sample data
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
                sheet = workbook.Worksheets[0];
            }
            else
            {
                workbook = new Workbook();
                sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("X");
                sheet.Cells["B1"].PutValue("Y");
                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue(i - 1);          // X values 1..5
                    sheet.Cells[$"B{i}"].PutValue((i - 1) * 2);   // Y values 2,4,6,8,10
                }
            }

            // Ensure at least one chart exists; create a simple line chart if none are present
            Chart chart;
            if (sheet.Charts.Count == 0)
            {
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries[0].XValues = "A2:A6";
            }
            else
            {
                chart = sheet.Charts[0];
            }

            // Add a linear trendline to the first series
            int trendlineIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIndex];
            trendline.Name = "LinearTrend";
            trendline.DisplayEquation = true; // Show equation on the chart (optional)

            // Aspose.Cells does not expose the trendline equation directly.
            // As a placeholder, we use the trendline name.
            string formula = trendline.Name ?? "Formula not available";

            // Save the workbook with the trendline
            string outputPath = "output_with_trendline.xlsx";
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }

            // Output the retrieved information to the console
            Console.WriteLine("Extracted Trendline Identifier:");
            Console.WriteLine(formula);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
