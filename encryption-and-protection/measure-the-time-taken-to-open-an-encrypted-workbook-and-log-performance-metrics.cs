// Title: Measure and log the time required to open a password‑protected Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an encrypted .xlsx file with Aspose.Cells, starts a Stopwatch, and prints the elapsed milliseconds and formatted TimeSpan after the workbook is opened. | Show how to set LoadOptions with a password, catch CellsException for an invalid password, and still report the time elapsed before the exception was thrown. | Demonstrate logging of both successful and failed workbook opening attempts, including performance metrics, using Console output in a .NET console application.
// Common Searches: how to benchmark opening speed of a password protected Excel file using Aspose.Cells in C# | c# measure time to load encrypted .xlsx with Aspose.Cells LoadOptions | catch invalid password error when loading protected workbook with Aspose.Cells and get elapsed time | performance logging for Aspose.Cells workbook load operation
// Tags: Aspose.Cells load encrypted workbook timing | C# Stopwatch Aspose.Cells performance | LoadOptions password Aspose.Cells | CellsException invalid password handling | benchmark workbook opening latency Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example verifies the encrypted Excel file exists, configures LoadOptions with the workbook password, and starts a Stopwatch. It then attempts to load the workbook using Aspose.Cells. On success, the timer stops and the elapsed time is printed in milliseconds and as a formatted TimeSpan. If a CellsException occurs, the code checks for an "Invalid password" message, reports the specific error, and logs the time elapsed before the failure. Any other exceptions are caught, the timer is stopped, and the error is displayed.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook and its password
        string filePath = "encrypted.xlsx";
        string password = "yourPassword";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        // Prepare load options with the password
        LoadOptions loadOptions = new LoadOptions
        {
            Password = password
        };

        // Start measuring time
        Stopwatch stopwatch = Stopwatch.StartNew();

        try
        {
            // Load the encrypted workbook
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Stop measuring time after successful load
            stopwatch.Stop();

            // Log performance metrics
            Console.WriteLine("Workbook opened successfully.");
            Console.WriteLine($"Time taken to open encrypted workbook: {stopwatch.ElapsedMilliseconds} ms");
            Console.WriteLine($"Time taken (hh:mm:ss.fff): {stopwatch.Elapsed}");
        }
        catch (CellsException ex)
        {
            stopwatch.Stop();

            // Check if the exception is due to an invalid password
            if (ex.Message != null && ex.Message.IndexOf("Invalid password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: Invalid password provided for the encrypted workbook.");
            }
            else
            {
                Console.WriteLine($"CellsException occurred: {ex.Message}");
            }

            Console.WriteLine($"Elapsed time before failure: {stopwatch.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
