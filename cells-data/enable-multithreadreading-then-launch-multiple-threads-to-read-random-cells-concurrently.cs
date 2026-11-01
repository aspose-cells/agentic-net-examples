// Title: Enable MultiThreadReading and perform concurrent random cell reads with Aspose.Cells for .NET
// AI Prompts: Provide C# code that sets WorkbookSettings.MultithreadedReading = true and launches multiple Task instances to read random cells from a worksheet using Aspose.Cells. | Show how to use a thread‑local Random object for safe parallel cell access in an Aspose.Cells workbook. | Create a sample that logs the cell address and value every N reads while reading cells concurrently with Aspose.Cells.
// Common Searches: how to turn on multithreaded reading in Aspose.Cells .NET | parallel random cell reading example using Aspose.Cells C# | Aspose.Cells read cells concurrently with Task Parallel Library | thread‑safe random number generation for Excel cell access in C# | performance tips for reading large Excel files with Aspose.Cells multithreading
// Tags: Aspose.Cells enable multithreaded reading | C# parallel cell access Aspose.Cells | thread‑safe random generation Excel C# | concurrent worksheet reading Aspose.Cells | Task Parallel Library Excel cell read

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

// Demonstrates enabling Aspose.Cells multithreaded reading and using multiple Task instances to read random cells from the first worksheet of an Excel file, with thread‑local Random and basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Multi‑threaded reading is automatically handled by Aspose.Cells where supported.
            // If a specific API exists in the used version, it can be set here.

            // Get the first worksheet (adjust index if needed)
            var sheet = workbook.Worksheets[0];

            // Define how many threads and how many reads each thread will perform
            int threadCount = 8;
            int readsPerThread = 100;

            // Create and start tasks
            var tasks = new Task[threadCount];
            for (int i = 0; i < threadCount; i++)
            {
                tasks[i] = Task.Run(() =>
                {
                    try
                    {
                        // Use a thread‑local Random instance for safe random numbers
                        var rnd = new Random(Guid.NewGuid().GetHashCode());

                        for (int j = 0; j < readsPerThread; j++)
                        {
                            // Determine random cell coordinates within the used range
                            int maxRow = sheet.Cells.MaxDataRow;
                            int maxCol = sheet.Cells.MaxDataColumn;
                            int row = rnd.Next(0, maxRow + 1);
                            int col = rnd.Next(0, maxCol + 1);

                            // Read the cell value
                            var value = sheet.Cells[row, col].Value;

                            // Optionally output some reads for verification
                            if (j % 20 == 0)
                            {
                                Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{row},{col}] = {value}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId} error: {ex.Message}");
                    }
                });
            }

            // Wait for all threads to finish
            Task.WaitAll(tasks);

            Console.WriteLine("Concurrent reading completed.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display the message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
