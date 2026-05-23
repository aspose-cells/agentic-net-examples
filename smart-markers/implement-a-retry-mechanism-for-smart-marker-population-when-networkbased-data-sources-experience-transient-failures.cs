using System;
using System.IO;
using System.Net;
using System.Threading;
using Aspose.Cells;

namespace SmartMarkerRetryDemo
{
    // Simple configuration holder for retry settings
    internal static class Config
    {
        public static int MaxRetryTimes { get; set; } = 3;
        public static int InitialFailRetryDelay { get; set; } = 2000; // milliseconds
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Configure retry policy for AI/network calls
                Config.MaxRetryTimes = 3;                 // Maximum number of retry attempts
                Config.InitialFailRetryDelay = 2000;      // Initial delay in milliseconds (2 seconds)

                // Create a new workbook and add smart markers
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("&=$Product.Name");
                sheet.Cells["A2"].PutValue("&=$Product.Price");

                // Initialize WorkbookDesigner with the workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Example JSON data source that could be fetched from a remote service
                // (In a real scenario this might involve an HTTP call that can fail)
                string jsonData = "{'Name':'Gadget','Price':99.99}";
                designer.SetJsonDataSource("Product", jsonData);

                // Populate smart markers with retry logic
                PopulateSmartMarkersWithRetry(designer);

                // Save the result (using the standard save lifecycle)
                string outputPath = "SmartMarkerOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <summary>
        /// Executes designer.Process() with a retry mechanism based on Config settings.
        /// </summary>
        /// <param name="designer">The WorkbookDesigner instance to process.</param>
        private static void PopulateSmartMarkersWithRetry(WorkbookDesigner designer)
        {
            int attempt = 0;
            while (true)
            {
                try
                {
                    // Attempt to process smart markers
                    designer.Process();
                    // Success – exit the loop
                    break;
                }
                catch (Exception ex) when (IsTransientFailure(ex))
                {
                    attempt++;
                    if (attempt > Config.MaxRetryTimes)
                    {
                        // Exceeded maximum retries; rethrow the exception
                        throw new InvalidOperationException(
                            $"Failed to populate smart markers after {Config.MaxRetryTimes} retries.", ex);
                    }

                    // Calculate delay (simple linear back‑off)
                    int delay = Config.InitialFailRetryDelay * attempt;
                    Console.WriteLine($"Transient failure detected (attempt {attempt}). Retrying in {delay} ms...");

                    // Wait before the next retry
                    Thread.Sleep(delay);
                }
            }
        }

        /// <summary>
        /// Determines whether an exception is likely caused by a transient network failure.
        /// </summary>
        /// <param name="ex">The exception to evaluate.</param>
        /// <returns>True if the exception is considered transient; otherwise false.</returns>
        private static bool IsTransientFailure(Exception ex)
        {
            // Common transient exceptions: WebException, TimeoutException, etc.
            return ex is WebException ||
                   ex is TimeoutException ||
                   (ex.InnerException != null && IsTransientFailure(ex.InnerException));
        }
    }
}