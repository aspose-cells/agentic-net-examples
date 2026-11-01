// Title: Generate a per‑worksheet theme accent color usage report in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells count how many cells use Accent1‑Accent6 theme colors in each worksheet | Create a summary sheet that lists theme accent color frequencies per worksheet with Aspose.Cells | How to iterate over cell styles to collect theme color statistics in an Excel file using Aspose.Cells | Aspose.Cells generate report of font, foreground and background theme accent usage across sheets | Save a new worksheet with theme color usage counts using Aspose.Cells .NET
// Tags: theme accent color counting with Aspose.Cells | worksheet level theme color statistics | add summary worksheet for accent usage | iterate cell styles for theme colors C# | export accent count report to new sheet | Aspose.Cells theme color analysis Excel

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ThemeColorUsageReport
{
    // The example loads an input .xlsx file with Aspose.Cells, scans every cell in each worksheet's used range, tallies occurrences of the six theme accent colors (Accent1‑Accent6) found in font, foreground, or background styles, writes the counts to a newly added worksheet named "ThemeColorUsageReport", and saves the workbook to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Prepare a list to hold usage data for each worksheet
                var usageData = new List<WorksheetUsage>();

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Initialize counters for the six theme accent colors
                    int[] accentCounts = new int[6]; // Index 0 => Accent1, 1 => Accent2, ...

                    // Determine the used range to limit iteration
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    // Scan each cell within the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            if (cell == null) continue;

                            Style style = cell.GetStyle();

                            // Check font theme color
                            if (IsAccentThemeColor(style.Font.ThemeColor, out int accentIndex))
                            {
                                accentCounts[accentIndex]++;
                            }

                            // Check foreground (pattern) theme color if set
                            if (IsAccentThemeColor(style.ForegroundThemeColor, out accentIndex))
                            {
                                accentCounts[accentIndex]++;
                            }

                            // Check background theme color if set
                            if (IsAccentThemeColor(style.BackgroundThemeColor, out accentIndex))
                            {
                                accentCounts[accentIndex]++;
                            }
                        }
                    }

                    // Store the results for this worksheet
                    usageData.Add(new WorksheetUsage
                    {
                        WorksheetName = sheet.Name,
                        AccentCounts = accentCounts
                    });
                }

                // Add a new worksheet to hold the report
                int reportIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportIndex];
                reportSheet.Name = "ThemeColorUsageReport";

                // Write header row
                reportSheet.Cells[0, 0].PutValue("Worksheet");
                for (int i = 0; i < 6; i++)
                {
                    reportSheet.Cells[0, i + 1].PutValue($"Accent{i + 1}");
                }

                // Write data rows
                for (int i = 0; i < usageData.Count; i++)
                {
                    WorksheetUsage data = usageData[i];
                    int row = i + 1;
                    reportSheet.Cells[row, 0].PutValue(data.WorksheetName);
                    for (int j = 0; j < 6; j++)
                    {
                        reportSheet.Cells[row, j + 1].PutValue(data.AccentCounts[j]);
                    }
                }

                // Save the workbook with the new report sheet
                workbook.Save(outputPath);
                Console.WriteLine($"Report saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to determine if a ThemeColor corresponds to an accent color
        private static bool IsAccentThemeColor(ThemeColor themeColor, out int accentIndex)
        {
            accentIndex = -1;

            string name = themeColor.ToString();

            // Theme colors for accents are named "Accent1" through "Accent6"
            if (name.StartsWith("Accent", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(name.Substring(6), out int num) && num >= 1 && num <= 6)
                {
                    accentIndex = num - 1; // zero‑based index
                    return true;
                }
            }

            return false;
        }

        // Simple container for worksheet usage information
        private class WorksheetUsage
        {
            public string WorksheetName { get; set; } = string.Empty;
            public int[] AccentCounts { get; set; } = new int[6];
        }
    }
}
