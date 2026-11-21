// Title: Add a horizontal stacked bar chart to an existing XLSX workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add horizontal stacked bar chart to an existing Excel file | How to insert a stacked bar chart into a loaded workbook with Aspose.Cells | Aspose.Cells ChartType.BarStacked example for .NET developers | Saving an XLSX workbook after adding a chart with Aspose.Cells C# | Load workbook, create stacked bar chart, and export to new file using Aspose.Cells
// Tags: Aspose.Cells stacked bar chart insertion | C# add chart to existing XLSX workbook | ChartType.BarStacked usage example | horizontal stacked bar chart Aspose.Cells | save modified workbook with chart Aspose | load workbook and create chart Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample verifies the input file, loads it into a Workbook, adds a horizontal stacked bar chart (ChartType.BarStacked) to the first worksheet, sets a title, defines a data series from range B2:B5, and saves the updated workbook as a new XLSX file while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Add a stacked bar (horizontal) chart
            // Parameters: chart type, upper‑left row, upper‑left column, lower‑right row, lower‑right column
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Stacked Bar Chart";

            // Add series values (B2:B5). Categories will be taken from the first column by default.
            int seriesIndex = chart.NSeries.Add("B2:B5", true);
            // Note: In recent Aspose.Cells versions, Series.CategoryData may not be available.
            // If needed, you can set categories via NSeries.Add with a separate range.

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Chart added and workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
