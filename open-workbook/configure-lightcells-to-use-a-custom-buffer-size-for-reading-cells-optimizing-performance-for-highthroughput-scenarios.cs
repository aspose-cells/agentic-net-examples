// Title: Aspose.Cells C# – LightCells with FileCache buffer and multithreaded reading
// Description: Shows how to implement a custom LightCellsDataHandler, set Cells.MemorySetting to FileCache, enable Cells.MultiThreadReading, and process rows in parallel to read large Excel workbooks efficiently with low memory consumption.
// Keywords: Aspose.Cells | LightCells | FileCache | MultiThreadReading | custom buffer | C# .NET | large workbook processing | parallel cell reading | streaming Excel | memory optimization
// Common Searches: Aspose.Cells LightCells custom buffer | set LightCells memory setting FileCache | enable multithreaded reading Aspose.Cells | process large Excel file with LightCells | C# LightCellsDataHandler example
// Developer Intent: Implement LightCells mode with a FileCache buffer and parallel reading to improve performance on massive Excel files.
// Use Cases: Read and transform a multi‑gigabyte Excel file without loading it fully into memory. | Log or analyze cell values in a streaming fashion using multiple threads. | Perform high‑throughput data extraction from spreadsheets in server‑side or background services.
// AI Prompts: Write C# code that sets up a LightCellsDataHandler, configures Cells.MemorySetting = MemorySetting.FileCache, and turns on Cells.MultiThreadReading for a workbook. | Explain how to tune the LightCells buffer size and thread count for optimal throughput in Aspose.Cells. | Provide a thread‑safe pattern for processing rows in parallel using LightCells mode.

using System;
using System.Threading;
using Aspose.Cells;

namespace LightCellsCustomBufferDemo
{
    // Custom handler for reading cells in LightCells mode
    // Shows how to implement a custom LightCellsDataHandler, set Cells.MemorySetting to FileCache, enable Cells.MultiThreadReading, and process rows in parallel to read large Excel workbooks efficiently with low memory consumption.
    public class CustomLightCellsDataHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet)
        {
            // Process all sheets
            return true;
        }

        public bool StartRow(int rowIndex)
        {
            // Process all rows
            return true;
        }

        public bool ProcessRow(Row row)
        {
            // No special row processing needed
            return true;
        }

        public bool StartCell(int columnIndex)
        {
            // Process all cells
            return true;
        }

        public bool ProcessCell(Cell cell)
        {
            // Example: simply output cell address and value
            Console.WriteLine($"Cell[{cell.Row},{cell.Column}] = {cell.Value}");
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to a large workbook to be processed
            string inputPath = "LargeWorkbook.xlsx";

            // Configure LightCells data handler
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new CustomLightCellsDataHandler();

            // Load workbook using LightCells mode
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet's cells
            Cells cells = workbook.Worksheets[0].Cells;

            // Set memory usage to FileCache (uses temporary files as buffer)
            cells.MemorySetting = MemorySetting.FileCache;

            // Enable multi‑thread reading for higher throughput
            cells.MultiThreadReading = true;

            // Example: read cells concurrently using multiple threads
            int totalRows = cells.MaxDataRow + 1;
            int threadCount = 4;
            int rowsPerThread = totalRows / threadCount;
            Thread[] threads = new Thread[threadCount];

            for (int t = 0; t < threadCount; t++)
            {
                int startRow = t * rowsPerThread;
                int endRow = (t == threadCount - 1) ? totalRows : startRow + rowsPerThread;

                threads[t] = new Thread(() =>
                {
                    for (int r = startRow; r < endRow; r++)
                    {
                        // Access cell values; MultiThreadReading must be true
                        var value = cells[r, 0].Value;
                        // Simulate processing
                        // Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Row {r} Value = {value}");
                    }
                });
                threads[t].Start();
            }

            // Wait for all threads to finish
            foreach (Thread thread in threads)
                thread.Join();

            // Save the processed workbook
            workbook.Save("ProcessedLargeWorkbook.xlsx");
        }
    }
}
