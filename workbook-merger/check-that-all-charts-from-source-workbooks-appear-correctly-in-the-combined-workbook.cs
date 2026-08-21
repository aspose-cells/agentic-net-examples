// Title: Verify Chart Preservation When Merging Excel Workbooks with Aspose.Cells for .NET
// Description: A C# example that loads multiple source Excel files, records the chart count on each worksheet, merges them into a single workbook using the Combine method, refreshes all data, and then validates that every original chart is present in the combined file before saving.
// Keywords: Aspose.Cells combine workbooks | C# chart verification after merge | Excel chart preservation Aspose | RefreshAll charts Aspose.Cells | merged workbook chart count
// Common Searches: how to keep charts when combining Excel files with Aspose.Cells | C# verify charts after workbook merge | Aspose.Cells Combine method chart loss | check chart count in merged workbook | refresh charts after merging Excel workbooks
// Developer Intent: Ensure that every chart from each source workbook remains intact after using Aspose.Cells' Combine method.
// Use Cases: Load each source workbook and capture the number of charts per worksheet. | Merge the workbooks into a single destination workbook with Combine and call RefreshAll. | Iterate through the combined worksheets, compare actual chart counts with the recorded values, and log any discrepancies before saving.
// AI Prompts: Generate C# code that logs missing or extra charts after merging multiple Excel workbooks with Aspose.Cells. | Create an NUnit test that asserts the chart count on each worksheet of a combined workbook matches the original counts. | Refactor the verification loop to use LINQ for comparing expected and actual chart counts.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsChartVerification
{
    // A C# example that loads multiple source Excel files, records the chart count on each worksheet, merges them into a single workbook using the Combine method, refreshes all data, and then validates that every original chart is present in the combined file before saving.
    class Program
    {
        static void Main()
        {
            // Paths to source workbooks (replace with actual file locations)
            string[] sourcePaths = { "Source1.xlsx", "Source2.xlsx", "Source3.xlsx" };

            // Create the destination workbook that will hold the combined result
            Workbook combinedWorkbook = new Workbook();

            // Store chart information from each source workbook for later verification
            var sourceChartInfo = new List<(int sheetIndex, int chartCount)>();

            // Load each source workbook, record its chart data, and combine it into the destination
            foreach (string path in sourcePaths)
            {
                // Load source workbook
                Workbook sourceWorkbook = new Workbook(path);

                // Record chart count per worksheet in the source workbook
                foreach (Worksheet ws in sourceWorkbook.Worksheets)
                {
                    int chartCount = ws.Charts.Count;
                    sourceChartInfo.Add((ws.Index, chartCount));
                }

                // Combine the source workbook into the destination workbook
                combinedWorkbook.Combine(sourceWorkbook);
            }

            // Refresh all pivot tables and charts in the combined workbook (ensures data is up‑to‑date)
            combinedWorkbook.Worksheets.RefreshAll();

            // Verify that charts from each source worksheet are present in the combined workbook
            bool allChartsPresent = true;
            int verificationIndex = 0; // Index to walk through sourceChartInfo

            foreach (Worksheet ws in combinedWorkbook.Worksheets)
            {
                // Skip worksheets that were originally empty (no source chart info)
                if (verificationIndex >= sourceChartInfo.Count)
                    break;

                var (sourceSheetIdx, expectedChartCount) = sourceChartInfo[verificationIndex];

                // The combined workbook preserves the original sheet order, so the indices should match
                if (ws.Index == sourceSheetIdx)
                {
                    int actualChartCount = ws.Charts.Count;
                    if (actualChartCount != expectedChartCount)
                    {
                        allChartsPresent = false;
                        Console.WriteLine($"Mismatch in worksheet '{ws.Name}' (Index {ws.Index}): " +
                                          $"expected {expectedChartCount} chart(s), found {actualChartCount}.");
                    }
                    else
                    {
                        Console.WriteLine($"Worksheet '{ws.Name}' (Index {ws.Index}) contains the expected " +
                                          $"{actualChartCount} chart(s).");
                    }

                    verificationIndex++;
                }
            }

            // Final result
            if (allChartsPresent)
                Console.WriteLine("All charts from source workbooks are present in the combined workbook.");
            else
                Console.WriteLine("Some charts are missing or mismatched in the combined workbook.");

            // Save the combined workbook
            combinedWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
