// Title: Merge multiple Excel workbooks and verify that all worksheet charts are retained using Aspose.Cells for .NET
// Common Searches: how to merge multiple .xlsx files and keep charts using Aspose.Cells C# | validate chart count after copying worksheets with Aspose.Cells in .NET | copy worksheets with embedded charts to a new workbook using Aspose.Cells | Aspose.Cells AddCopy method preserving charts when combining workbooks | check for missing charts after merging Excel workbooks with Aspose.Cells
// Tags: Aspose.Cells workbook merging preserving charts | AddCopy method copy worksheet with charts | chart integrity check Aspose.Cells | combine Excel files chart retention .NET | verify merged workbook charts C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads a list of source .xlsx files, records the number of charts on each worksheet, copies each worksheet into a new workbook using Workbook.Worksheets.AddCopy, then iterates the combined workbook to compare actual chart counts with the recorded values, reporting any mismatches before saving the merged workbook.
class ChartCombiner
{
    static void Main()
    {
        try
        {
            // List of source workbook file paths to combine
            List<string> sourceFiles = new List<string>
            {
                "Source1.xlsx",
                "Source2.xlsx",
                "Source3.xlsx"
            };

            // Create a new empty workbook that will hold the combined content
            Workbook combinedWorkbook = new Workbook();

            // Remove the default empty worksheet created by the constructor (if any)
            if (combinedWorkbook.Worksheets.Count > 0)
            {
                combinedWorkbook.Worksheets.RemoveAt(0);
            }

            // Keep a dictionary to track chart counts per worksheet name for verification
            Dictionary<string, int> sourceChartCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (string filePath in sourceFiles)
            {
                // Verify that the source file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}. Skipping.");
                    continue;
                }

                // Load each source workbook
                Workbook sourceWorkbook = new Workbook(filePath);

                // Iterate through all worksheets in the source workbook
                foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
                {
                    // Record the number of charts in the source worksheet
                    int chartCount = sourceSheet.Charts.Count;
                    sourceChartCounts[sourceSheet.Name] = chartCount;

                    // Copy the worksheet (including its charts) into the combined workbook
                    // AddCopy expects the worksheet name, not the Worksheet object
                    combinedWorkbook.Worksheets.AddCopy(sourceSheet.Name);
                }
            }

            // Verification: ensure each copied worksheet contains the same number of charts as the source
            foreach (Worksheet combinedSheet in combinedWorkbook.Worksheets)
            {
                string sheetName = combinedSheet.Name;

                if (sourceChartCounts.TryGetValue(sheetName, out int expectedChartCount))
                {
                    int actualChartCount = combinedSheet.Charts.Count;

                    if (actualChartCount != expectedChartCount)
                    {
                        Console.WriteLine($"Chart mismatch in worksheet '{sheetName}': expected {expectedChartCount}, found {actualChartCount}.");
                    }
                    else
                    {
                        Console.WriteLine($"Worksheet '{sheetName}' contains all {actualChartCount} chart(s) as expected.");
                    }
                }
            }

            // Save the combined workbook to a new file
            string outputPath = "CombinedWorkbook.xlsx";
            combinedWorkbook.Save(outputPath);
            Console.WriteLine($"Combined workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
