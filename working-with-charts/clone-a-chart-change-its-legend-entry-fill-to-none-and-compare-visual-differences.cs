// Title: Clone an Excel chart, modify the first legend entry’s font color, and programmatically compare legend colors with Aspose.Cells for .NET
// AI Prompts: Duplicate the first chart on a worksheet, set the first legend entry’s font color to black, and save the workbook using Aspose.Cells in C#. | Write C# code that reads the font colors of the first legend entries from the original and cloned charts and outputs whether they are identical. | Create a new chart area (rows 20‑30, columns 0‑10) by cloning an existing chart and copy its data range with Aspose.Cells.
// Common Searches: Aspose.Cells C# clone chart and change legend entry font color | How to compare legend entry colors of two Excel charts using Aspose.Cells | Programmatically duplicate a chart and edit its legend in .NET | Detect visual differences between original and cloned charts with Aspose.Cells | Set legend entry font color in an Aspose.Cells chart example
// Tags: chart cloning with Aspose.Cells .NET | set legend entry font color Aspose.Cells C# | compare legend colors Aspose.Cells chart | duplicate Excel chart programmatically | visual difference detection in Excel charts Aspose

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, clones the first chart to a new location, copies its data range, changes the first legend entry’s font color in the cloned chart, compares the font colors of the first legend entries between the original and cloned charts, reports any differences, and saves the modified workbook.
class ChartCloneAndCompare
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing the original chart
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the worksheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            Chart originalChart = sheet.Charts[0];

            // Clone the chart to a new position (rows 20‑30, columns 0‑10)
            int clonedChartIndex = sheet.Charts.Add(originalChart.Type, 20, 0, 30, 10);
            Chart clonedChart = sheet.Charts[clonedChartIndex];

            // Copy the data range from the original chart to the cloned chart
            string dataRange = originalChart.GetChartDataRange();
            if (!string.IsNullOrEmpty(dataRange))
            {
                // The second argument indicates whether the data is plotted vertically.
                clonedChart.SetChartDataRange(dataRange, true);
            }

            // Change the font color of the first legend entry of the cloned chart
            if (clonedChart.Legend != null && clonedChart.Legend.LegendEntries.Count > 0)
            {
                LegendEntry legendEntry = clonedChart.Legend.LegendEntries[0];
                legendEntry.Font.Color = Color.Black; // Example modification
            }

            // Simple comparison: check if the first legend entry font colors are identical
            bool areIdentical = true;
            int diffCount = 0;

            if (originalChart.Legend != null && originalChart.Legend.LegendEntries.Count > 0 &&
                clonedChart.Legend != null && clonedChart.Legend.LegendEntries.Count > 0)
            {
                Color originalColor = originalChart.Legend.LegendEntries[0].Font.Color;
                Color clonedColor = clonedChart.Legend.LegendEntries[0].Font.Color;

                if (originalColor.ToArgb() != clonedColor.ToArgb())
                {
                    areIdentical = false;
                    diffCount = 1;
                }
            }

            // Output the comparison result
            Console.WriteLine(areIdentical
                ? "The charts are visually identical (based on legend font color)."
                : $"The charts differ in {diffCount} property(ies).");

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the cloned chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
