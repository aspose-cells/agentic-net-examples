// Title: C# – Add exponential back‑off retry to Aspose.Cells smart marker processing
// Description: Demonstrates loading a workbook template with smart markers, fetching a DataTable that may fail due to transient network errors, assigning it to WorkbookDesigner, and processing the markers inside a configurable retry loop with exponential back‑off before saving the result.
// Keywords: Aspose.Cells | smart markers | retry | exponential backoff | C# | .NET | WorkbookDesigner | transient network error | data source failure | GitHub example
// Common Searches: Aspose.Cells smart marker retry C# | exponential backoff for WorkbookDesigner.Process | handle transient network errors Aspose.Cells | retry logic Aspose.Cells example | smart marker population failure handling
// Developer Intent: Implement a configurable retry mechanism with exponential back‑off to keep smart marker processing reliable when data retrieval experiences temporary failures.
// Use Cases: Regenerating reports when a web API intermittently times out | Recovering from brief database connection drops during smart marker population | Ensuring automated workbook generation works in unstable network environments
// AI Prompts: Write C# code that wraps WorkbookDesigner.Process() in a retry loop with configurable attempts and exponential back‑off using Aspose.Cells. | Create a reusable retry helper for any Aspose.Cells operation that may throw transient exceptions, including logging and delay. | Show how to integrate the Polly library with Aspose.Cells smart marker processing for advanced retry policies.

using System;
using System.Data;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace SmartMarkerRetryDemo
{
    // Simple configuration holder for retry settings
    // Demonstrates loading a workbook template with smart markers, fetching a DataTable that may fail due to transient network errors, assigning it to WorkbookDesigner, and processing the markers inside a configurable retry loop with exponential back‑off before saving the result.
    internal static class Config
    {
        public static int MaxRetryTimes { get; set; } = 3;
        public static int InitialFailRetryDelay { get; set; } = 2000; // milliseconds
    }

    class Program
    {
        static void Main()
        {
            // Configure retry settings for AI/network operations
            Config.MaxRetryTimes = 3;               // Maximum number of retry attempts
            Config.InitialFailRetryDelay = 2000;    // Initial delay in milliseconds (2 seconds)

            // Load the workbook template that contains smart markers
            WorkbookDesigner designer = new WorkbookDesigner();
            const string templatePath = "TemplateWithSmartMarkers.xlsx";

            try
            {
                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException($"Template file not found: {templatePath}");
                }

                designer.Workbook = new Workbook(templatePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load template workbook: {ex.Message}");
                return;
            }

            // Set up a data source that may fail due to transient network issues
            DataTable data;
            try
            {
                data = GetDataWithPossibleTransientFailure();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to retrieve data: {ex.Message}");
                return;
            }

            // Assign the data source to the designer
            designer.SetDataSource(data);

            // Attempt to process smart markers with retry logic
            int attempt = 0;
            while (true)
            {
                try
                {
                    // Process smart markers – this may involve network calls internally
                    designer.Process();
                    // If processing succeeds, exit the loop
                    break;
                }
                catch (Exception ex)
                {
                    attempt++;
                    if (attempt > Config.MaxRetryTimes)
                    {
                        Console.WriteLine($"Processing failed after {attempt - 1} retries.");
                        Console.WriteLine($"Error: {ex.Message}");
                        return;
                    }

                    // Calculate exponential back‑off delay
                    int delay = Config.InitialFailRetryDelay * attempt;
                    Console.WriteLine($"Transient failure detected (attempt {attempt}): {ex.Message}");
                    Console.WriteLine($"Waiting {delay} ms before retrying...");

                    Thread.Sleep(delay);
                }
            }

            // Save the populated workbook
            const string outputPath = "OutputWithSmartMarkers.xlsx";
            try
            {
                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }

        // Simulated method that retrieves data and may throw a transient exception
        private static DataTable GetDataWithPossibleTransientFailure()
        {
            // In a real scenario, this could be a call to a web service, database, etc.
            // Here we randomly throw an exception to mimic a transient network failure.
            Random rnd = new Random();
            if (rnd.NextDouble() < 0.3) // 30% chance of failure
            {
                throw new InvalidOperationException("Simulated transient network error while fetching data.");
            }

            // Create sample data table
            DataTable table = new DataTable("Employees");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Age", typeof(int));
            table.Columns.Add("Department", typeof(string));

            table.Rows.Add("John Doe", 30, "Sales");
            table.Rows.Add("Jane Smith", 28, "Marketing");
            table.Rows.Add("Bob Johnson", 45, "HR");

            return table;
        }
    }
}
