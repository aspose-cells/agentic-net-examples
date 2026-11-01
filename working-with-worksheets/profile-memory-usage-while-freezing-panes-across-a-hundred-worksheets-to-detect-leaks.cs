// Title: Measure .NET memory consumption when applying FreezePanes to 100 worksheets with Aspose.Cells
// AI Prompts: Write a C# console program that creates a workbook, adds 100 worksheets, calls FreezePanes(1,1,0,0) on each sheet, and logs GC.GetTotalMemory after every iteration. | Enhance the profiling code to capture the workbook's byte size after each FreezePanes call and export the measurements to a CSV file for later analysis. | Extend the example to save the workbook in both XLSX and CSV formats, record memory usage before and after each save, and compare the memory impact of the two formats.
// Common Searches: how to detect memory leaks in Aspose.Cells when freezing panes on many worksheets | C# profiling memory usage of a workbook with 100 sheets using GC.GetTotalMemory | benchmark freeze panes performance across multiple worksheets in Aspose.Cells | measure memory growth after applying FreezePanes in a large Aspose.Cells workbook | Aspose.Cells memory consumption when adding and freezing rows and columns
// Tags: memory profiling Aspose.Cells FreezePanes | GC.GetTotalMemory workbook creation .NET | detect memory leak adding worksheets Aspose.Cells | freeze panes performance measurement | Aspose.Cells workbook size after save

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryProfiling
{
    // The sample creates a new Workbook, adds 100 worksheets, names each sheet, applies FreezePanes(1,1,0,0), writes a value to A1, forces garbage collection, records memory usage with GC.GetTotalMemory after each sheet, saves the workbook to a MemoryStream as XLSX, and prints per‑worksheet memory deltas and total memory increase to help identify potential leaks.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // List to hold memory usage after each worksheet processing
                List<long> memoryUsage = new List<long>();

                // Force garbage collection before starting measurement
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // Record baseline memory usage
                long baselineMemory = GC.GetTotalMemory(true);
                Console.WriteLine($"Baseline memory: {baselineMemory:N0} bytes");

                // Create a new workbook (allowed creation rule)
                Workbook workbook = new Workbook();

                // Process 100 worksheets
                for (int i = 0; i < 100; i++)
                {
                    Worksheet sheet;
                    if (i == 0)
                    {
                        // First worksheet already exists in a new workbook
                        sheet = workbook.Worksheets[0];
                    }
                    else
                    {
                        // Add additional worksheet and retrieve it by index (Add returns int in some versions)
                        int newIndex = workbook.Worksheets.Add();
                        sheet = workbook.Worksheets[newIndex];
                    }

                    sheet.Name = $"Sheet{i + 1}";

                    // Apply Freeze Panes (freeze first row and first column)
                    // Use the overload that accepts four parameters for compatibility
                    sheet.FreezePanes(1, 1, 0, 0);

                    // Optional: write some data to ensure the sheet has content
                    sheet.Cells["A1"].PutValue($"Data in {sheet.Name}");

                    // Force a collection to get a more accurate measurement per iteration
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();

                    // Record memory usage after processing this worksheet
                    long currentMemory = GC.GetTotalMemory(true);
                    memoryUsage.Add(currentMemory);
                    Console.WriteLine($"After processing {sheet.Name}: {currentMemory:N0} bytes (Δ {currentMemory - baselineMemory:N0})");
                }

                // Save the workbook to a memory stream (allowed save rule)
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to memory stream. Size: {ms.Length:N0} bytes");
                }

                // Analyze memory usage differences
                Console.WriteLine("\nMemory usage delta per worksheet:");
                for (int i = 0; i < memoryUsage.Count; i++)
                {
                    long delta = i == 0 ? memoryUsage[i] - baselineMemory : memoryUsage[i] - memoryUsage[i - 1];
                    Console.WriteLine($"Worksheet {i + 1}: Δ {delta:N0} bytes");
                }

                // Final memory snapshot
                long finalMemory = GC.GetTotalMemory(true);
                Console.WriteLine($"\nFinal memory after all operations: {finalMemory:N0} bytes");
                Console.WriteLine($"Total increase from baseline: {finalMemory - baselineMemory:N0} bytes");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
