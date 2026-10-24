// Title: Benchmark loading time and memory usage of a password‑protected XLSX workbook using Aspose.Cells for .NET
// Common Searches: how to measure load time of an encrypted Excel file with Aspose.Cells in C# | C# code to log memory consumption when opening a password protected workbook using Aspose.Cells | benchmark performance of Aspose.Cells LoadOptions for encrypted XLSX files | record elapsed milliseconds while loading a protected workbook with Aspose.Cells .NET | detect invalid password error when timing workbook load in Aspose.Cells
// Tags: Aspose.Cells encrypted workbook load timing | memory profiling Aspose.Cells LoadOptions | benchmarking password protected XLSX loading .NET | handle CellsException invalid password Aspose.Cells | performance metrics for Aspose.Cells workbook opening

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample verifies the encrypted Excel file exists, captures memory usage before loading, starts a Stopwatch, and opens the workbook with LoadOptions that include the password. After loading, it stops the timer, records memory usage again, and prints elapsed milliseconds, memory before, memory after, and additional memory allocated. It also catches CellsException to identify an invalid password and ensures the stopwatch is stopped for any exception.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook and its password
        string workbookPath = "encrypted.xlsx";
        string password = "myPassword";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        // Capture memory usage before loading
        long memoryBefore = GC.GetTotalMemory(forceFullCollection: true);

        // Start timing
        Stopwatch stopwatch = Stopwatch.StartNew();

        try
        {
            // Set load options with the password for the encrypted file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Open the encrypted workbook
            Workbook workbook = new Workbook(workbookPath, loadOptions);

            // Stop timing after successful load
            stopwatch.Stop();

            // Capture memory usage after loading
            long memoryAfter = GC.GetTotalMemory(forceFullCollection: true);

            // Log performance metrics
            Console.WriteLine($"Time taken to open encrypted workbook: {stopwatch.ElapsedMilliseconds} ms");
            Console.WriteLine($"Memory used before load: {memoryBefore / 1024} KB");
            Console.WriteLine($"Memory used after load: {memoryAfter / 1024} KB");
            Console.WriteLine($"Additional memory allocated: {(memoryAfter - memoryBefore) / 1024} KB");
        }
        catch (CellsException ex)
        {
            stopwatch.Stop();
            // Determine if the exception is due to an invalid password
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: Invalid password provided for the encrypted workbook.");
            }
            else
            {
                Console.WriteLine($"CellsException occurred: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
