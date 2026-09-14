// Title: How to profile memory usage while freezing panes on 100 worksheets with Aspose.Cells for .NET
// AI Prompts: Write a C# console program that creates a workbook, adds 100 worksheets, applies FreezePanes(1,1,1,1) to each sheet, forces garbage collection, and logs the memory usage after each sheet using GC.GetTotalMemory. | Extend the program to export the recorded memory values and per‑sheet deltas to a CSV file and calculate the average memory increase per worksheet.
// Common Searches: Aspose.Cells memory leak detection when using FreezePanes on many sheets | C# measure memory consumption after applying FreezePanes to each worksheet | profile .NET Excel workbook memory while adding and freezing 100 worksheets | log GC.GetTotalMemory for each worksheet in Aspose.Cells
// Tags: Aspose.Cells worksheet freeze panes performance | C# Excel workbook memory analysis | GC.GetTotalMemory profiling Aspose.Cells | detect Excel memory leaks .NET | benchmark worksheet freeze operation

using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Cells;

// The sample creates a workbook, adds 100 worksheets, freezes the first row and column on each sheet, forces garbage collection after each operation, records total memory using GC.GetTotalMemory, and prints per‑sheet memory differences to help identify potential leaks.
class FreezePanesMemoryProfile
{
    static void Main()
    {
        try
        {
            // List to store memory usage after each worksheet operation
            List<long> memoryUsage = new List<long>();

            // Force a full garbage collection before starting the test
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Record baseline memory usage
            long baselineMemory = GC.GetTotalMemory(true);
            Console.WriteLine($"Baseline memory: {baselineMemory} bytes");

            // Create a new workbook (contains a default worksheet)
            Workbook workbook = new Workbook();

            // Rename the default worksheet to avoid name clash with added sheets
            workbook.Worksheets[0].Name = "Sheet0";

            // Loop to create 100 worksheets and freeze panes on each
            for (int i = 0; i < 100; i++)
            {
                // Add a new worksheet with a unique name
                string sheetName = $"Sheet{i + 1}";
                Worksheet newSheet = workbook.Worksheets.Add(sheetName);

                // Freeze the first row and first column (equivalent to FreezePanes at row 2, column 2)
                newSheet.FreezePanes(1, 1, 1, 1);

                // Optional: add some data to make the sheet realistic
                newSheet.Cells["A1"].PutValue("Header");
                newSheet.Cells["B2"].PutValue(i);

                // Force garbage collection to get a more accurate measurement
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // Capture memory usage after processing this worksheet
                long currentMemory = GC.GetTotalMemory(true);
                memoryUsage.Add(currentMemory);
                Console.WriteLine($"After sheet {i + 1}: {currentMemory} bytes (Δ {currentMemory - baselineMemory} bytes)");
            }

            // Analyze memory usage to detect potential leaks
            Console.WriteLine("\nMemory usage differences between consecutive sheets:");
            for (int i = 1; i < memoryUsage.Count; i++)
            {
                long diff = memoryUsage[i] - memoryUsage[i - 1];
                Console.WriteLine($"Sheet {i} -> Sheet {i + 1}: Δ {diff} bytes");
            }

            // Save the workbook (optional, not required for profiling)
            // workbook.Save("FreezePanesTest.xlsx");

            Console.WriteLine("\nProfiling completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
