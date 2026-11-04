// Title: Benchmark the load time of a password‑protected XLSX workbook with Aspose.Cells for .NET
// Common Searches: how to measure load time of encrypted Excel file using Aspose.Cells C# | Aspose.Cells load encrypted workbook performance timing example | C# Stopwatch for opening password protected XLSX with Aspose.Cells | log elapsed milliseconds when loading protected workbook Aspose.Cells | detect invalid password while timing workbook load Aspose.Cells .NET
// Tags: Aspose.Cells encrypted workbook load timing | C# Stopwatch load performance Aspose.Cells | LoadOptions password protected Excel loading | benchmark Excel decryption with Aspose.Cells | measure workbook opening latency .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample checks for the existence of an encrypted.xlsx file, starts a Stopwatch, loads the workbook with a password via LoadOptions, stops the timer, and prints the elapsed milliseconds and worksheet count. It catches CellsException to identify invalid passwords, logs the time elapsed before the failure, and handles any other exceptions with appropriate messages.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook file
        string filePath = "encrypted.xlsx";

        // Password required to open the workbook
        string password = "myPassword";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found – {filePath}");
            return;
        }

        // Start measuring time
        Stopwatch stopwatch = Stopwatch.StartNew();

        try
        {
            // Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Stop measuring time
            stopwatch.Stop();

            // Log performance metrics
            Console.WriteLine($"Time taken to open encrypted workbook: {stopwatch.ElapsedMilliseconds} ms");
            Console.WriteLine($"Number of worksheets loaded: {workbook.Worksheets.Count}");
        }
        catch (CellsException ex)
        {
            stopwatch.Stop();

            // Check if the exception message indicates an invalid password
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: Invalid password provided for the workbook.");
            }
            else
            {
                Console.WriteLine($"CellsException: {ex.Message}");
            }

            Console.WriteLine($"Elapsed time before failure: {stopwatch.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
