// Title: Concurrent random cell reading with multiple threads using Aspose.Cells for .NET
// Common Searches: how to read random cells from an Excel file using multiple threads in C# Aspose.Cells | Aspose.Cells multithreaded cell access example .NET | parallel reading of Excel cells with Aspose.Cells and Task.Run | enable thread‑safe random cell extraction from a workbook using Aspose.Cells | C# code to read Excel cells concurrently with Aspose.Cells library
// Tags: multithreaded cell reading Aspose.Cells .NET | parallel Excel cell access C# | random cell extraction using Aspose.Cells | thread‑safe worksheet reading Aspose.Cells | Task.Run Excel cell read Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, then starts ten parallel tasks that each generate a random cell address (A1‑Z100), read the cell value from the first worksheet, and write the result to the console in a thread‑safe way. It demonstrates Aspose.Cells' built‑in multithreaded reading support and proper synchronization for random cell access.
class MultiThreadCellReader
{
    static void Main()
    {
        try
        {
            const string filePath = "Sample.xlsx";

            // Verify that the workbook file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // NOTE: Multi‑threaded reading is enabled by default in recent Aspose.Cells versions.
            // If a specific setting is required, use the appropriate API for the version you target.

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Thread‑safe random number generator
            Random globalRand = new Random();
            object randLock = new object();

            int threadCount = 10;
            List<Task> tasks = new List<Task>();

            for (int i = 0; i < threadCount; i++)
            {
                tasks.Add(Task.Run(() =>
                {
                    try
                    {
                        // Generate a random cell address (A1 to Z100)
                        int row, col;
                        lock (randLock)
                        {
                            row = globalRand.Next(0, 100);   // rows 0‑99 (Excel rows 1‑100)
                            col = globalRand.Next(0, 26);    // columns 0‑25 (A‑Z)
                        }

                        // Read the cell value
                        Cell cell = sheet.Cells[row, col];
                        string address = cell.Name; // e.g., "B5"
                        object value = cell.Value;

                        // Thread‑safe console output
                        lock (Console.Out)
                        {
                            Console.WriteLine($"Thread {Task.CurrentId}: Cell {address} = {value ?? "null"}");
                        }
                    }
                    catch (Exception ex)
                    {
                        lock (Console.Out)
                        {
                            Console.WriteLine($"Thread {Task.CurrentId} encountered an error: {ex.Message}");
                        }
                    }
                }));
            }

            // Wait for all tasks to finish
            Task.WaitAll(tasks.ToArray());
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
