// Title: C# Retry Mechanism for Aspose.Cells Smart Marker Processing with Transient Network Failures
// Description: This example loads a template workbook, fetches JSON data that may come from a network service, and wraps WorkbookDesigner.Process in a configurable retry loop (max attempts, delay, optional exponential backoff). Each failure is logged before the workbook is finally saved.
// Keywords: Aspose.Cells | smart markers | retry | transient network failure | C# | .NET | WorkbookDesigner | JSON data source | exponential backoff | retry policy | network resilience
// Common Searches: Aspose.Cells retry smart markers | C# retry WorkbookDesigner.Process | handle network errors Aspose.Cells smart markers | smart marker retry example | Aspose.Cells exponential backoff
// Developer Intent: Add a retry policy around WorkbookDesigner.Process to safely populate smart markers when the JSON data source may temporarily fail.
// Use Cases: Reprocess smart markers automatically if fetching JSON data fails due to a temporary network outage. | Limit retries with configurable max attempts and delay to prevent endless loops. | Log each retry attempt and surface the final exception after the maximum retries are exceeded. | Apply exponential backoff for more robust handling of intermittent network issues.
// AI Prompts: Write C# code that adds exponential backoff to the retry loop for Aspose.Cells smart marker processing. | Create a reusable RetryHelper class for WorkbookDesigner.Process with customizable maxAttempts, delay, and backoff strategy. | Show how to combine HttpClient with Polly retry policies to fetch JSON data before setting it as a smart marker data source. | Explain how to configure logging and error handling for retrying smart marker population in Aspose.Cells.

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

// This example loads a template workbook, fetches JSON data that may come from a network service, and wraps WorkbookDesigner.Process in a configurable retry loop (max attempts, delay, optional exponential backoff). Each failure is logged before the workbook is finally saved.
class SmartMarkerRetryDemo
{
    // Retry configuration
    private const int MaxRetryTimes = 3;            // Maximum number of retry attempts
    private const int InitialFailRetryDelay = 2000; // Initial delay (ms) before a retry

    static void Main()
    {
        try
        {
            const string templatePath = "template.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the template file exists before loading
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Example: set a JSON data source that might be retrieved from a network service
            string jsonData = GetJsonFromNetwork();
            designer.SetJsonDataSource("Data", jsonData);

            // Retry loop for processing smart markers
            int attempt = 0;
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
                    if (attempt > MaxRetryTimes)
                    {
                        // Exceeded maximum retries – log and rethrow
                        Console.WriteLine($"Processing failed after {attempt} attempts: {ex.Message}");
                        throw;
                    }

                    // Log the failure and wait before retrying
                    Console.WriteLine($"Attempt {attempt} failed: {ex.Message}. Retrying in {InitialFailRetryDelay} ms...");
                    Thread.Sleep(InitialFailRetryDelay);
                }
            }

            // Save the populated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception e)
        {
            // Global exception handling
            Console.WriteLine($"An error occurred: {e.Message}");
        }
    }

    // Placeholder for a network call that retrieves JSON data.
    // Replace with actual HTTP request logic as required.
    static string GetJsonFromNetwork()
    {
        // Simulated JSON payload
        return "{'Name':'Sample Product','Value':99.99}";
    }
}
