using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryProfiling
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (creation rule)
                Workbook workbook = new Workbook();

                // Optional: use a memory‑saving mode to reduce pressure during the test
                workbook.Settings.MemorySetting = MemorySetting.FileCache;

                // Get the current process to read memory usage
                Process proc = Process.GetCurrentProcess();

                // Record initial memory usage
                long initialMemory = proc.PrivateMemorySize64;
                Console.WriteLine($"Initial memory: {initialMemory / 1024 / 1024} MB");

                // Ensure there is at least one worksheet to start with
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Freeze panes on the first sheet
                sheet.FreezePanes("C3", 3, 3);
                Console.WriteLine($"After freezing Sheet1: {proc.PrivateMemorySize64 / 1024 / 1024} MB");

                // Add and process 99 additional worksheets (total 100)
                for (int i = 2; i <= 100; i++)
                {
                    // Add a new worksheet (creation rule) – Worksheets.Add() returns the index
                    int newIndex = workbook.Worksheets.Add();
                    Worksheet ws = workbook.Worksheets[newIndex];
                    ws.Name = $"Sheet{i}";

                    // Freeze panes at cell C3 with 3 frozen rows and 3 frozen columns
                    ws.FreezePanes("C3", 3, 3);

                    // Capture memory after each freeze to spot any abnormal growth
                    long currentMemory = proc.PrivateMemorySize64;
                    Console.WriteLine($"After freezing {ws.Name}: {currentMemory / 1024 / 1024} MB");
                }

                // Final memory usage
                long finalMemory = proc.PrivateMemorySize64;
                Console.WriteLine($"Final memory: {finalMemory / 1024 / 1024} MB");
                Console.WriteLine($"Total increase: {(finalMemory - initialMemory) / 1024 / 1024} MB");

                // Save the workbook (save rule)
                string outputPath = "MemoryProfileFreezePanes.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}