// Title: C# Memory Profiling of FreezePanes Across 100 Worksheets with Aspose.Cells
// Description: A console example that creates a workbook, optionally switches to FileCache mode, adds 100 worksheets, fills each with sample data, freezes panes at cell C3, forces garbage collection, records the process's private memory after each FreezePanes call, prints the results, and saves the file.
// Keywords: Aspose.Cells | C# | memory profiling | FreezePanes | Workbook performance | FileCache mode | private memory measurement | leak detection | process memory usage
// Common Searches: Aspose.Cells memory leak detection | profile memory when freezing panes in C# | measure memory usage per worksheet Aspose.Cells | use FileCache to reduce memory pressure Aspose.Cells | how to track private memory size in .NET
// Developer Intent: Track memory consumption after each FreezePanes operation to spot growth patterns or leaks in large workbooks.
// Use Cases: Generate a per‑sheet memory report for performance tuning. | Compare FileCache versus default in‑memory settings. | Validate that FreezePanes does not cause unmanaged memory growth.
// AI Prompts: Show C# code that logs Process.PrivateMemorySize64 before and after Worksheet.FreezePanes for every sheet in a workbook. | Explain why invoking GC.Collect and GC.WaitForPendingFinalizers improves memory measurement accuracy in Aspose.Cells profiling. | Suggest additional Aspose.Cells settings or disposal patterns to minimize memory usage when freezing panes on many worksheets.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace FreezePanesMemoryProfiling
{
    // A console example that creates a workbook, optionally switches to FileCache mode, adds 100 worksheets, fills each with sample data, freezes panes at cell C3, forces garbage collection, records the process's private memory after each FreezePanes call, prints the results, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                using (Workbook workbook = new Workbook())
                {
                    // Use FileCache mode to reduce memory pressure (optional)
                    workbook.Settings.MemorySetting = MemorySetting.FileCache;

                    // List to store memory usage after each freeze operation
                    List<long> memoryUsages = new List<long>();

                    // Process for memory measurement
                    Process currentProcess = Process.GetCurrentProcess();

                    // Freeze panes on 100 worksheets
                    for (int i = 0; i < 100; i++)
                    {
                        Worksheet sheet;

                        // First worksheet already exists; subsequent ones are added
                        if (i == 0)
                        {
                            sheet = workbook.Worksheets[0];
                        }
                        else
                        {
                            // Add returns the index of the new sheet; retrieve the worksheet by that index
                            int newIndex = workbook.Worksheets.Add();
                            sheet = workbook.Worksheets[newIndex];
                        }

                        // Give the worksheet a distinct name
                        sheet.Name = $"Sheet{i + 1}";

                        // Populate some sample data (optional, to simulate realistic usage)
                        for (int row = 0; row < 20; row++)
                        {
                            for (int col = 0; col < 10; col++)
                            {
                                sheet.Cells[row, col].PutValue($"R{row}C{col}");
                            }
                        }

                        // Freeze panes at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns
                        sheet.FreezePanes(2, 2, 2, 2);

                        // Force garbage collection for a more accurate measurement
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();

                        // Record current private memory size (in bytes)
                        currentProcess.Refresh();
                        memoryUsages.Add(currentProcess.PrivateMemorySize64);
                    }

                    // Output memory usage after each freeze operation
                    Console.WriteLine("Memory usage after each FreezePanes call (bytes):");
                    for (int i = 0; i < memoryUsages.Count; i++)
                    {
                        Console.WriteLine($"Worksheet {i + 1}: {memoryUsages[i]}");
                    }

                    // Determine output path and ensure directory exists
                    string outputPath = "FreezePanesMemoryProfile.xlsx";
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the workbook
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }

                Console.WriteLine("Resources disposed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
