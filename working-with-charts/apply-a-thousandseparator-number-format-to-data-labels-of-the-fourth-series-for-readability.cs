// Title: Apply a thousand‑separator number format to the data labels of the fourth series in an Excel chart using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing workbook, retrieves the first chart, enables data labels for the series at index 3, and sets its NumberFormat to "#,##0". | Show how to verify a chart contains at least four series before applying a thousand‑separator format to the fourth series' data labels with Aspose.Cells. | Modify the example to apply the "#,##0" number format to data labels of every series in a chart using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# format data labels with thousand separator for a specific chart series | How to set NumberFormat "#,##0" on the fourth series data labels in an Excel chart using Aspose.Cells | C# example for enabling data labels and applying custom number format to chart series in Aspose.Cells | Apply thousand separator to chart series data labels in a workbook with Aspose.Cells .NET
// Tags: set DataLabels.NumberFormat for specific chart series Aspose.Cells | enable data labels for fourth series Aspose.Cells C# | apply thousand separator number format chart data labels .NET | format chart series data labels Aspose.Cells | verify chart series count before formatting Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads a workbook, accesses the first worksheet's first chart, enables data labels for the fourth series, applies the "#,##0" thousand‑separator number format to those labels, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Ensure the chart has a fourth series (index 3)
            if (chart.NSeries.Count <= 3)
            {
                Console.WriteLine("The chart does not contain a fourth series.");
                return;
            }

            Series fourthSeries = chart.NSeries[3];

            // Show data labels for the fourth series
            // Use the correct property name for the current Aspose.Cells version
            fourthSeries.DataLabels.ShowValue = true;

            // Apply thousand‑separator number format to the data labels
            fourthSeries.DataLabels.NumberFormat = "#,##0";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
