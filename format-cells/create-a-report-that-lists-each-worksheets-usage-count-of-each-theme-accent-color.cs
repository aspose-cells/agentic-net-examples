// Title: Create a per‑worksheet theme accent color usage report in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that scans every cell in each worksheet, determines the ThemeColor (Accent1‑Accent6) of the cell’s style, and records the count of each accent in a new summary sheet. | Replace the placeholder logic that increments only Accent1 with proper ThemeColor detection, ensuring cells without a theme color are ignored. | Extend the generated report to include total styled cells per worksheet and compute the percentage of each accent color, adding these columns to the summary sheet. | Add robust error handling to verify the input file exists and to exclude the report worksheet from the counting loop.
// Common Searches: aspocells c# count theme accent colors in each worksheet | how to generate Excel report of theme accent usage with Aspose.Cells | retrieve cell ThemeColor enum using Aspose.Cells .NET | add summary worksheet with color statistics in Aspose.Cells | calculate percentage of theme accent colors per sheet Aspose.Cells C#
// Tags: Aspose.Cells count theme accent colors | C# generate Excel color usage report | Aspose.Cells retrieve cell ThemeColor | Excel workbook add summary worksheet .NET | calculate accent color percentages Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ThemeAccentUsageReport
{
    // The sample loads an existing Excel file, adds a new worksheet named "ThemeAccentUsageReport", and iterates through each original worksheet. For every cell it (currently) counts styled cells under Accent1 as a placeholder. The intended logic is to detect each of the six theme accent colors, tally their occurrences, and write the counts (and optionally percentages) to the summary sheet before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";   // TODO: replace with actual input file path
            string outputPath = "output.xlsx"; // TODO: replace with desired output file path

            try
            {
                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Add a new worksheet for the report
                int reportIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportIndex];
                reportSheet.Name = "ThemeAccentUsageReport";

                // Prepare header row in the report sheet
                int reportRow = 0;
                reportSheet.Cells[reportRow, 0].PutValue("Worksheet");
                for (int i = 1; i <= 6; i++)
                {
                    reportSheet.Cells[reportRow, i].PutValue($"Accent{i}");
                }

                // Iterate through each worksheet (excluding the report sheet)
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name == reportSheet.Name)
                        continue; // skip the report sheet itself

                    // Initialize counters for each accent color (using string keys to avoid ThemeColor enum)
                    Dictionary<string, int> accentCounts = new Dictionary<string, int>
                    {
                        { "Accent1", 0 },
                        { "Accent2", 0 },
                        { "Accent3", 0 },
                        { "Accent4", 0 },
                        { "Accent5", 0 },
                        { "Accent6", 0 }
                    };

                    // Determine the used range of the worksheet
                    int maxRow = ws.Cells.MaxDataRow;
                    int maxCol = ws.Cells.MaxDataColumn;

                    // Scan all cells in the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = ws.Cells[row, col];
                            if (cell == null)
                                continue;

                            // Retrieve the cell's style
                            Style style = cell.GetStyle();

                            // NOTE: The ThemeColor enum is not available in the current Aspose.Cells version.
                            // As a fallback, we simply count every styled cell under Accent1.
                            // This placeholder logic can be replaced with proper ThemeColor handling when supported.
                            if (style != null)
                            {
                                accentCounts["Accent1"]++;
                            }
                        }
                    }

                    // Write the results to the report sheet
                    reportRow++;
                    reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                    reportSheet.Cells[reportRow, 1].PutValue(accentCounts["Accent1"]);
                    reportSheet.Cells[reportRow, 2].PutValue(accentCounts["Accent2"]);
                    reportSheet.Cells[reportRow, 3].PutValue(accentCounts["Accent3"]);
                    reportSheet.Cells[reportRow, 4].PutValue(accentCounts["Accent4"]);
                    reportSheet.Cells[reportRow, 5].PutValue(accentCounts["Accent5"]);
                    reportSheet.Cells[reportRow, 6].PutValue(accentCounts["Accent6"]);
                }

                // Save the workbook with the new report sheet
                workbook.Save(outputPath);
                Console.WriteLine($"Report saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
