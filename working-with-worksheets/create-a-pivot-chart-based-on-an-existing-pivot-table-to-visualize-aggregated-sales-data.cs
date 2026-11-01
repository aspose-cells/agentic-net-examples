// Title: Generate a column pivot chart from an existing pivot table in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to programmatically add a pivot chart to an Excel worksheet with Aspose.Cells C# | Aspose.Cells create column chart based on pivot table data in .NET | link pivot table to chart series using Aspose.Cells API | set chart title and move legend to the right with Aspose.Cells | save an Excel file after adding a pivot chart using Aspose.Cells for .NET
// Tags: Aspose.Cells create pivot chart | C# column chart from pivot table | Aspose.Cells link chart series to pivot data | Excel workbook save with added chart Aspose.Cells | Aspose.Cells set chart legend position

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The example loads an existing Excel workbook, checks for a pivot table, adds a column chart linked to that pivot table, customizes the chart title and legend placement, and saves the updated file.
class PivotChartExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load workbook containing a pivot table
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure at least one pivot table is present
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 1, 5, 20, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Aggregated Sales";

            // Link chart to pivot table data
            // The first parameter is the source data name (pivot table name), second indicates whether to use categories
            chart.NSeries.Add(pivotTable.Name, true);

            // Position legend on the right
            chart.Legend.Position = LegendPositionType.Right;

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
