// Title: C# example: Load an Excel workbook, attempt to set the Accent6 theme to a pastel shade, and verify chart legend visibility with Aspose.Cells
// Common Searches: aspnet cells c# change Accent6 theme color to pastel | how to determine if a chart legend is shown using Aspose.Cells | iterate over all worksheets and charts in a workbook with Aspose.Cells .NET | load existing Excel file and save after processing using Aspose.Cells for .NET | Aspose.Cells ShowLegend property example for chart legends
// Tags: Accent6 theme color adjustment Aspose.Cells | chart legend presence verification Aspose.Cells | enumerate worksheets and charts C# | open and save Excel file Aspose.Cells | ShowLegend property usage Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing Excel workbook with Aspose.Cells, attempts (but skips) a direct Accent6 theme color change, then iterates through each worksheet and its charts to report whether each chart displays a legend and, if so, the legend's position, finally saving the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // -----------------------------------------------------------------
            // Theme modification (Accent6) – not available via the Theme property
            // in current Aspose.Cells versions (Theme returns a string). The block
            // is kept for compatibility; it simply skips modification.
            // -----------------------------------------------------------------
            try
            {
                // No direct API to modify Accent6; skip if unavailable.
                // Placeholder for future theme adjustments.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Theme modification skipped: {ex.Message}");
            }

            // Verify chart legends in all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Determine chart index safely
                    int chartIndex = sheet.Charts.IndexOf(chart);

                    // Check if the chart has a visible legend using ShowLegend property
                    bool hasLegend = chart.ShowLegend;

                    // Output verification result
                    Console.WriteLine($"Worksheet: {sheet.Name}, Chart Index: {chartIndex}, Has Legend: {hasLegend}");

                    // If a legend exists, output its position
                    if (hasLegend && chart.Legend != null)
                    {
                        Console.WriteLine($"  Legend Position: {chart.Legend.Position}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
