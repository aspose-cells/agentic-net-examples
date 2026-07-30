// Title: Measure and log opening time of an encrypted Excel workbook with Aspose.Cells for .NET
// Description: C# sample that verifies an encrypted Excel file, applies LoadOptions with a password, times the Workbook constructor using Stopwatch, outputs elapsed milliseconds and the workbook's IsEncrypted flag, and handles invalid‑password and other exceptions.
// Keywords: Aspose.Cells encrypted workbook load time | measure Excel file opening performance .NET | benchmark password‑protected workbook Aspose | Stopwatch load options Aspose.Cells | C# log workbook opening duration
// Common Searches: how to time opening an encrypted Excel file with Aspose.Cells | Aspose.Cells .NET measure load time of password protected workbook | log elapsed milliseconds when loading encrypted workbook C# | check IsEncrypted property after opening workbook Aspose | performance test for opening protected Excel using Aspose.Cells
// Developer Intent: The developer wants to benchmark the duration required to open a password‑protected Excel workbook and verify its encryption status.
// Use Cases: Profile load performance of large, password‑protected workbooks in an automated ETL pipeline. | Record the IsEncrypted flag during batch imports to enforce security compliance. | Add opening‑time metrics to CI/CD tests to catch regressions after upgrading Aspose.Cells.
// AI Prompts: Create C# code that measures the time to open a password‑protected Excel file with Aspose.Cells and writes the result to a log file. | Show how to run the opening operation repeatedly and compute average, minimum, and maximum load times for more stable benchmarking. | Explain how to extend error handling to distinguish wrong password, unsupported encryption algorithm, and corrupted file while still capturing load duration.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // C# sample that verifies an encrypted Excel file, applies LoadOptions with a password, times the Workbook constructor using Stopwatch, outputs elapsed milliseconds and the workbook's IsEncrypted flag, and handles invalid‑password and other exceptions.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted workbook and its password
            string filePath = "encrypted.xlsx";
            string password = "password";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Configure load options with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            Workbook workbook = null;
            Stopwatch stopwatch = new Stopwatch();

            try
            {
                // Start timing the workbook opening operation
                stopwatch.Start();

                // Open the encrypted workbook using the constructor that accepts LoadOptions
                workbook = new Workbook(filePath, loadOptions);

                // Stop timing
                stopwatch.Stop();

                // Log performance metrics
                Console.WriteLine($"Time to open encrypted workbook: {stopwatch.ElapsedMilliseconds} ms");
                Console.WriteLine($"Workbook IsEncrypted property: {workbook.Settings.IsEncrypted}");
            }
            catch (CellsException ex)
            {
                stopwatch.Stop();

                // Check if the exception message indicates an invalid password
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("Error: Invalid password provided for the encrypted workbook.");
                }
                else
                {
                    Console.WriteLine($"Aspose.Cells error: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
            finally
            {
                // Clean up
                workbook?.Dispose();
            }
        }
    }
}
