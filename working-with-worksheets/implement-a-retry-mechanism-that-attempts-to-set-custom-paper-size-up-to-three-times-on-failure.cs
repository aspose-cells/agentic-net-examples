// Title: C# Retry Logic for Aspose.Cells PageSetup.CustomPaperSize (max 3 attempts)
// Description: Shows how to set a 1.5" × 2.0" custom paper size on an Aspose.Cells worksheet using a retry loop that tries up to three times, logs each failure, and saves the workbook.
// Keywords: Aspose.Cells | C# | PageSetup.CustomPaperSize | retry loop | exception handling | custom paper size | workbook save | transient error handling | Polly resilience | automation
// Common Searches: Aspose.Cells retry setting custom paper size | C# PageSetup.CustomPaperSize exception handling | how to retry Aspose.Cells page setup | set custom paper size with retry in C# | Aspose.Cells transient error retry
// Developer Intent: Add a robust retry mechanism that attempts to set a custom paper size up to three times, handling any exceptions before saving the workbook.
// Use Cases: Ensuring a specific page dimension when a printer or service intermittently rejects the size. | Automated report generation that must enforce exact paper dimensions despite occasional API hiccups. | Batch processing of many workbooks where each file’s page setup needs reliable retry to avoid stopping the whole job.
// AI Prompts: Generate a reusable C# method that wraps PageSetup.CustomPaperSize with configurable max attempts and optional delay. | Write unit tests that mock exceptions from CustomPaperSize and verify the retry loop succeeds after a transient failure. | Refactor the retry code to use the Polly library for exponential back‑off when setting custom paper size in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsRetryDemo
{
    // Shows how to set a 1.5" × 2.0" custom paper size on an Aspose.Cells worksheet using a retry loop that tries up to three times, logs each failure, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object
            PageSetup pageSetup = worksheet.PageSetup;

            // Desired custom paper size in inches
            double widthInches = 1.5;
            double heightInches = 2.0;

            // Maximum number of attempts
            int maxAttempts = 3;
            int attempt = 0;
            bool success = false;

            // Retry loop for setting custom paper size
            while (attempt < maxAttempts && !success)
            {
                try
                {
                    attempt++;
                    // Attempt to set the custom paper size (method rule)
                    pageSetup.CustomPaperSize(widthInches, heightInches);
                    success = true; // If no exception, operation succeeded
                }
                catch (Exception ex)
                {
                    // Log the failure and continue to next attempt
                    Console.WriteLine($"Attempt {attempt} failed: {ex.Message}");

                    if (attempt >= maxAttempts)
                    {
                        // All attempts exhausted; rethrow or handle as needed
                        Console.WriteLine("All retry attempts exhausted. Unable to set custom paper size.");
                        throw;
                    }
                }
            }

            // Save the workbook (save rule)
            workbook.Save("CustomPaperSizeWithRetry.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
    }
}
