// Title: Combine multiple Excel workbooks into one and ensure all charts are copied correctly using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads several .xlsx files, copies each worksheet with its charts into a new Workbook using Aspose.Cells, and saves the merged file. | Add logic to count the charts in each source worksheet, compare those counts with the corresponding worksheets in the merged workbook, and output any mismatches. | Include robust error handling for missing source files and exceptions during loading or saving, and log verification results to the console.
// Common Searches: aspnet c# combine multiple Excel files keep charts Aspose.Cells | how to verify chart count after merging workbooks with Aspose.Cells | copy worksheets with embedded charts using Aspose.Cells AddCopy method | merge Excel workbooks and detect missing charts in C# | handle missing source workbook files when merging with Aspose.Cells
// Tags: addcopy worksheet charts Aspose.Cells | chart count verification Aspose.Cells | merge multiple workbooks C# Aspose.Cells | handle missing source files Aspose.Cells | combined workbook chart integrity

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ChartCombinationDemo
{
    // The sample loads a list of source .xlsx files, copies each worksheet (including its charts) into a new workbook using Workbook.Worksheets.AddCopy, records chart counts for source and destination worksheets, compares them to detect mismatches, handles missing files and other errors, and saves the combined workbook as CombinedWorkbook.xlsx.
    class Program
    {
        static void Main()
        {
            // List of source workbook file paths
            List<string> sourceFiles = new List<string>
            {
                "SourceWorkbook1.xlsx",
                "SourceWorkbook2.xlsx",
                // add more source files as needed
            };

            // ---------- Create destination workbook ----------
            Workbook combinedWorkbook = new Workbook();

            // Remove the default empty worksheet that Aspose.Cells creates (if any)
            if (combinedWorkbook.Worksheets.Count > 0)
                combinedWorkbook.Worksheets.RemoveAt(0);

            // Keep track of chart counts for verification
            Dictionary<string, int> sourceChartCounts = new Dictionary<string, int>();
            Dictionary<string, int> combinedChartCounts = new Dictionary<string, int>();

            // ---------- Load each source workbook and copy its worksheets ----------
            foreach (string sourcePath in sourceFiles)
            {
                try
                {
                    // Ensure the source file exists before attempting to load
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"Source file not found: {sourcePath}. Skipping.");
                        continue;
                    }

                    // Load source workbook
                    Workbook sourceWorkbook = new Workbook(sourcePath);

                    // Record chart counts per worksheet in the source workbook
                    foreach (Worksheet srcWs in sourceWorkbook.Worksheets)
                    {
                        string key = $"{sourcePath}:{srcWs.Name}";
                        sourceChartCounts[key] = srcWs.Charts.Count;
                    }

                    // Copy each worksheet (including its charts) to the combined workbook
                    foreach (Worksheet srcWs in sourceWorkbook.Worksheets)
                    {
                        // AddCopy expects the worksheet name, not the Worksheet object
                        int newIndex = combinedWorkbook.Worksheets.AddCopy(srcWs.Name);

                        // Record chart counts for the newly added worksheet
                        Worksheet destWs = combinedWorkbook.Worksheets[newIndex];
                        string destKey = $"{sourcePath}:{srcWs.Name}";
                        combinedChartCounts[destKey] = destWs.Charts.Count;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{sourcePath}': {ex.Message}");
                }
            }

            // ---------- Verification ----------
            // Ensure that every chart present in the source worksheets appears in the combined workbook
            foreach (var kvp in sourceChartCounts)
            {
                string worksheetKey = kvp.Key;
                int sourceCount = kvp.Value;
                int combinedCount = combinedChartCounts.ContainsKey(worksheetKey) ? combinedChartCounts[worksheetKey] : -1;

                if (sourceCount != combinedCount)
                {
                    Console.WriteLine($"Chart mismatch in worksheet '{worksheetKey}'. Source charts: {sourceCount}, Combined charts: {combinedCount}");
                }
                else
                {
                    Console.WriteLine($"Charts verified for worksheet '{worksheetKey}'. Count: {sourceCount}");
                }
            }

            // ---------- Save the combined workbook ----------
            try
            {
                combinedWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Combined workbook saved successfully as 'CombinedWorkbook.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving combined workbook: {ex.Message}");
            }
        }
    }
}
