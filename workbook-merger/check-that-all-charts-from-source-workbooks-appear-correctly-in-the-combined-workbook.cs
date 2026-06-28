using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsChartCombineCheck
{
    class Program
    {
        static void Main()
        {
            // Paths to source workbooks (replace with actual file paths)
            string[] sourceFiles = { "SourceWorkbook1.xlsx", "SourceWorkbook2.xlsx" };

            // Create an empty destination workbook
            Workbook destWorkbook = new Workbook();

            // Keep track of total number of charts in all source workbooks
            int totalSourceCharts = 0;

            // Iterate through each source workbook
            foreach (string sourcePath in sourceFiles)
            {
                // Load the source workbook
                Workbook srcWorkbook = new Workbook(sourcePath);

                // Count charts in the current source workbook
                for (int sheetIdx = 0; sheetIdx < srcWorkbook.Worksheets.Count; sheetIdx++)
                {
                    Worksheet srcSheet = srcWorkbook.Worksheets[sheetIdx];
                    totalSourceCharts += srcSheet.Charts.Count;
                }

                // Combine the source workbook into the destination workbook
                destWorkbook.Combine(srcWorkbook);
            }

            // Refresh all pivot tables and charts after combination
            destWorkbook.Worksheets.RefreshAll();

            // Count charts in the combined (destination) workbook
            int totalDestCharts = 0;
            for (int sheetIdx = 0; sheetIdx < destWorkbook.Worksheets.Count; sheetIdx++)
            {
                Worksheet destSheet = destWorkbook.Worksheets[sheetIdx];
                totalDestCharts += destSheet.Charts.Count;
            }

            // Verify that all source charts are present in the combined workbook
            if (totalDestCharts == totalSourceCharts)
            {
                Console.WriteLine("Success: All charts from source workbooks are present in the combined workbook.");
            }
            else
            {
                Console.WriteLine($"Mismatch: Source charts = {totalSourceCharts}, Combined workbook charts = {totalDestCharts}");
            }

            // Save the combined workbook
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}