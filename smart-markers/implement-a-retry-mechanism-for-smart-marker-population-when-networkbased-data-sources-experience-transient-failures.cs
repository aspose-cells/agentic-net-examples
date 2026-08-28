// Title: Implement configurable retry with exponential back‑off for Aspose.Cells smart marker processing in C# using a network‑based data source
// AI Prompts: Wrap WorkbookDesigner.Process() in a retry loop that applies exponential back‑off for transient network errors. | Create a reusable settings class to define maximum retry attempts and initial delay, and integrate it with the smart marker demo. | Enhance the example to read retry parameters from an external JSON configuration file and log each retry attempt with timestamps.
// Common Searches: how to add retry logic to Aspose.Cells smart marker processing in C# | C# Aspose.Cells WorkbookDesigner.Process transient network failure handling | exponential backoff strategy for smart markers population Aspose.Cells | configurable retry settings for Aspose.Cells smart markers in .NET
// Tags: WorkbookDesigner retry processing | smart marker backoff strategy | Aspose.Cells transient network retry | configurable retry settings .NET | smart marker data source resilience

using System;
using System.Collections;
using System.Threading;
using Aspose.Cells;

namespace SmartMarkerRetryDemo
{
    // Simple configuration holder for retry settings
    // Demonstrates configuring max retry attempts and initial delay, simulating a flaky network data source, and processing smart markers with a linear then exponential back‑off retry loop before saving the workbook.
    internal static class Config
    {
        public static int MaxRetryTimes { get; set; } = 3;
        public static int InitialFailRetryDelay { get; set; } = 500; // milliseconds
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Configure retry settings (optional – can be set elsewhere in the application)
                Config.MaxRetryTimes = DefaultMaxRetryTimes;          // maximum retry attempts
                Config.InitialFailRetryDelay = DefaultInitialDelayMs; // initial delay in ms

                // -----------------------------------------------------------------
                // Create a workbook with smart markers (template)
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Example smart markers that will be populated from a data source
                sheet.Cells["A1"].PutValue("&=Employees.Name");
                sheet.Cells["B1"].PutValue("&=Employees.Age");

                // Name the range that contains smart markers (required when using range smart markers)
                sheet.Cells.CreateRange("A1:B1").Name = "_CellsSmartMarkers";

                // -----------------------------------------------------------------
                // Prepare a data source that simulates a network call
                // -----------------------------------------------------------------
                ArrayList dataSource;
                try
                {
                    dataSource = GetNetworkBasedDataSource();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to obtain data source: {ex.Message}");
                    return;
                }

                // -----------------------------------------------------------------
                // Set up the WorkbookDesigner
                // -----------------------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                    // LineByLine is obsolete; range smart markers are used by default
                };

                // Assign the data source to the designer
                designer.SetDataSource("Employees", dataSource);

                // -----------------------------------------------------------------
                // Process smart markers with retry logic
                // -----------------------------------------------------------------
                ProcessSmartMarkersWithRetry(designer);

                // -----------------------------------------------------------------
                // Save the resulting workbook
                // -----------------------------------------------------------------
                workbook.Save("SmartMarkerResult.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        // Default retry configuration constants
        private const int DefaultMaxRetryTimes = 3;
        private const int DefaultInitialDelayMs = 500;

        /// <param name="designer">The WorkbookDesigner instance to process.</param>
        private static void ProcessSmartMarkersWithRetry(WorkbookDesigner designer)
        {
            int attempt = 0;
            int maxAttempts = Config.MaxRetryTimes > 0 ? Config.MaxRetryTimes : 1;
            int delay = Config.InitialFailRetryDelay > 0 ? Config.InitialFailRetryDelay : 0;

            while (true)
            {
                try
                {
                    // Attempt to process smart markers
                    designer.Process();
                    // If processing succeeds, exit the loop
                    break;
                }
                catch (Exception ex)
                {
                    attempt++;

                    // If we have exhausted all retry attempts, rethrow the exception
                    if (attempt > maxAttempts)
                    {
                        Console.WriteLine($"Smart marker processing failed after {attempt - 1} retries.");
                        throw;
                    }

                    // Log the transient failure and wait before retrying
                    Console.WriteLine($"Transient failure detected (attempt {attempt}/{maxAttempts}): {ex.Message}");
                    Console.WriteLine($"Waiting {delay} ms before next retry...");

                    // Wait for the configured delay (simple linear back‑off)
                    Thread.Sleep(delay);

                    // Optionally increase the delay for the next attempt (exponential back‑off)
                    delay *= 2;
                }
            }
        }

        /// <summary>
        /// Simulates a network‑based data source retrieval.
        /// </summary>
        /// <returns>An ArrayList containing employee objects.</returns>
        private static ArrayList GetNetworkBasedDataSource()
        {
            // Simulate a 30% chance of transient network failure
            Random rnd = new Random();
            if (rnd.NextDouble() < 0.3)
            {
                throw new InvalidOperationException("Simulated network timeout.");
            }

            // In a real implementation, replace the following with actual network data retrieval
            var list = new ArrayList
            {
                new Employee { Name = "Alice", Age = 30 },
                new Employee { Name = "Bob", Age = 45 }
            };

            return list;
        }

        // Simple POCO representing an employee (used as a data source)
        public class Employee
        {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; }
        }
    }
}
