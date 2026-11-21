// Title: C# example to profile memory usage while applying Freeze Panes to 100 worksheets using Aspose.Cells
// Common Searches: how to detect memory leaks when applying freeze panes to many worksheets with Aspose.Cells in C# | C# code to profile memory consumption of an Aspose.Cells workbook after each FreezePanes call | measure memory usage per worksheet while adding freeze panes in Aspose.Cells | performance testing of FreezePanes across 100 sheets using Aspose.Cells .NET | using GC.GetTotalMemory to monitor Aspose.Cells workbook memory during pane freezing
// Tags: freeze panes memory profiling Aspose.Cells | Aspose.Cells workbook memory leak detection | C# garbage collection memory measurement Aspose.Cells | processing 100 worksheets performance Aspose.Cells | measure per‑sheet memory usage Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace FreezePanesMemoryProfile
{
    // The sample creates a Workbook, iterates through 100 worksheets, adds or renames each sheet, applies FreezePanes to the first row and column, inserts simple data, forces garbage collection, records total memory after each sheet with GC.GetTotalMemory, outputs the differences to identify potential leaks, and finally saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // List to hold memory usage after each worksheet processing
                List<long> memoryUsage = new List<long>();

                // Force a full garbage collection before starting the test
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // Record initial memory usage
                long initialMemory = GC.GetTotalMemory(true);
                Console.WriteLine($"Initial memory: {initialMemory} bytes");

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Process 100 worksheets
                for (int i = 0; i < 100; i++)
                {
                    try
                    {
                        Worksheet sheet;

                        // The first worksheet already exists in the workbook
                        if (i == 0)
                        {
                            sheet = workbook.Worksheets[0];
                            sheet.Name = $"Sheet{i + 1}";
                        }
                        else
                        {
                            // Add a new worksheet and obtain the Worksheet object directly
                            sheet = workbook.Worksheets.Add($"Sheet{i + 1}");
                        }

                        // Apply Freeze Panes (freeze first row and first column)
                        // Use the overload that specifies total rows and columns to avoid version issues
                        sheet.FreezePanes(1, 1, sheet.Cells.MaxDataRow + 1, sheet.Cells.MaxDataColumn + 1);

                        // Optionally add some data to make the sheet realistic
                        sheet.Cells["A1"].PutValue("Header");
                        sheet.Cells["B2"].PutValue(i);

                        // Force garbage collection to get a more accurate measurement
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();

                        // Record memory usage after processing this worksheet
                        long currentMemory = GC.GetTotalMemory(true);
                        memoryUsage.Add(currentMemory);
                        Console.WriteLine($"After processing Sheet{i + 1}: {currentMemory} bytes (Δ {currentMemory - initialMemory} bytes)");
                    }
                    catch (Exception sheetEx)
                    {
                        Console.WriteLine($"Error processing sheet {i + 1}: {sheetEx.Message}");
                    }
                }

                // Analyze memory usage to detect potential leaks
                Console.WriteLine("\nMemory usage differences between consecutive sheets:");
                for (int i = 1; i < memoryUsage.Count; i++)
                {
                    long diff = memoryUsage[i] - memoryUsage[i - 1];
                    Console.WriteLine($"Sheet {i + 1} - Sheet {i}: {diff} bytes");
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                string outputPath = "FreezePanesTest.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to '{Path.GetFullPath(outputPath)}'.");

                Console.WriteLine("\nProcessing complete. Press any key to exit.");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
