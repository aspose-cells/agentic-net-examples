// Title: C# example: Load an XLSX workbook, add a linear trendline to the first chart series with Aspose.Cells, retrieve its equation, and save the file
// Common Searches: C# Aspose.Cells add linear trendline to first chart series in an existing XLSX workbook | How to retrieve the equation of a trendline from an Excel chart using Aspose.Cells .NET | Aspose.Cells recalculate chart after adding a trendline to access the equation label | Save a modified workbook that includes a new trendline with Aspose.Cells for .NET | Read trendline data label text from a chart series using Aspose.Cells C#
// Tags: Aspose.Cells add linear trendline to chart series | Aspose.Cells retrieve trendline equation C# | Aspose.Cells recalculate chart after trendline addition | Aspose.Cells save workbook with updated chart | Aspose.Cells chart series manipulation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an XLSX workbook, finds the first chart, adds a linear trendline to its first series, enables equation display, recalculates the chart to generate the label, extracts the equation text, logs it, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "input.xlsx";
        string outputPath = "output_with_trendline.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Locate the first chart in the workbook (search worksheets sequentially)
            Chart chart = null;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Charts.Count > 0)
                {
                    chart = ws.Charts[0];
                    break;
                }
            }

            if (chart == null)
            {
                Console.WriteLine("Error: No chart found in the workbook.");
                return;
            }

            // Ensure the chart has at least one series to which a trendline can be added
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("Error: The found chart does not contain any series.");
                return;
            }

            // Add a linear trendline to the first series of the chart
            int trendlineIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIndex];

            // Configure the trendline to display its equation on the chart
            trendline.DisplayEquation = true;
            trendline.DisplayRSquared = false; // optional: hide R‑squared value

            // Recalculate the chart so that the data label (which holds the equation) is generated
            chart.Calculate();

            // Retrieve the equation text from the trendline's data label
            string equation = trendline.DataLabels.Text;

            // Log the retrieved equation
            Console.WriteLine($"Trendline equation: {equation}");

            // Save the modified workbook to a new file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Modified workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An exception occurred: {ex.Message}");
        }
    }
}
