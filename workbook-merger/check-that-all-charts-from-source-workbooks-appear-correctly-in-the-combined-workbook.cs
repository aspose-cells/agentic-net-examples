// Title: Check that all charts from multiple source workbooks remain after using Workbook.Combine in Aspose.Cells for .NET
// AI Prompts: Load each source workbook, count its charts, combine them with Workbook.Combine, then assert that the combined workbook’s chart count equals the sum of the source counts. | Create C# code that merges several Excel files containing charts using Aspose.Cells and automatically verifies chart integrity by comparing pre‑merge and post‑merge chart totals. | Add exception handling that throws a descriptive error when the chart count after Workbook.Combine does not match the expected total.
// Common Searches: Aspose.Cells how to ensure charts are kept when merging multiple Excel files in C# | C# verify chart count after Workbook.Combine operation | compare chart totals before and after combining workbooks with Aspose.Cells | detect missing charts after merging Excel workbooks using Aspose.Cells .NET
// Tags: Workbook.Combine chart preservation | Aspose.Cells chart count validation | merge Excel workbooks with charts .NET | chart integrity after workbook combine | verify combined workbook chart total

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace ChartCombineVerification
{
    // The example loads a list of source Excel workbooks, counts the charts in each worksheet, merges them into a new workbook using Workbook.Combine, compares the total chart count before and after the merge, reports whether all charts are present, and saves the combined file as CombinedWorkbook.xlsx.
    class Program
    {
        static void Main()
        {
            // Paths to source workbooks that contain charts
            List<string> sourcePaths = new List<string>
            {
                "SourceWorkbook1.xlsx",
                "SourceWorkbook2.xlsx"
                // Add more paths as needed
            };

            // Create an empty destination workbook
            Workbook destWorkbook = new Workbook();

            // Keep track of total number of charts in all source workbooks
            int totalSourceCharts = 0;

            // Load each source workbook, count its charts, and combine it into the destination workbook
            foreach (string path in sourcePaths)
            {
                // Load source workbook
                Workbook sourceWorkbook = new Workbook(path);

                // Count charts in the current source workbook
                int sourceCharts = sourceWorkbook.Worksheets
                    .Cast<Worksheet>()
                    .Sum(ws => ws.Charts.Count);
                totalSourceCharts += sourceCharts;

                // Combine the source workbook into the destination workbook
                destWorkbook.Combine(sourceWorkbook);
            }

            // After combining, count charts in the combined workbook
            int combinedCharts = destWorkbook.Worksheets
                .Cast<Worksheet>()
                .Sum(ws => ws.Charts.Count);

            // Verify that all charts from source workbooks are present in the combined workbook
            if (combinedCharts == totalSourceCharts)
            {
                Console.WriteLine("All charts are present after combination.");
                Console.WriteLine($"Total source charts: {totalSourceCharts}, Combined workbook charts: {combinedCharts}");
            }
            else
            {
                Console.WriteLine("Chart count mismatch after combination.");
                Console.WriteLine($"Total source charts: {totalSourceCharts}, Combined workbook charts: {combinedCharts}");
            }

            // Save the combined workbook
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
