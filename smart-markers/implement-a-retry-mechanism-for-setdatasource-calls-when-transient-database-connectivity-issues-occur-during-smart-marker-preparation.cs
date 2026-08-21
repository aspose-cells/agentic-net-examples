// Title: C# Retry Wrapper for WorkbookDesigner.SetDataSource to Handle Transient DB Errors in Aspose.Cells Smart Markers
// Description: Demonstrates how to add configurable retry logic around WorkbookDesigner.SetDataSource using a helper method, enabling resilient smart‑marker processing when temporary database connectivity failures occur.
// Keywords: Aspose.Cells | WorkbookDesigner | SetDataSource | retry logic | transient database error | smart markers | C# | ExecuteWithRetry | exponential backoff | data source resilience | .NET Excel automation
// Common Searches: Aspose.Cells retry SetDataSource | smart markers retry mechanism C# | handle transient DB errors Aspose.Cells | WorkbookDesigner retry example | C# retry wrapper for Excel smart markers
// Developer Intent: Implement a configurable retry pattern around WorkbookDesigner.SetDataSource so the workbook generation continues despite intermittent database connectivity problems.
// Use Cases: Automatically re‑execute SetDataSource when a SqlException or network glitch occurs. | Control maximum attempts via a configuration setting. | Log each retry and apply incremental delay to give the database time to recover. | Integrate the retry helper into existing smart‑marker workflows without altering template logic.
// AI Prompts: Write C# code that adds exponential backoff with jitter to ExecuteWithRetry for Aspose.Cells SetDataSource. | Create a mock test that throws a transient SqlException on the first two SetDataSource calls and verifies three total attempts. | Generate Markdown documentation showing how to configure MaxRetryTimes and customize delay for smart‑marker data sources. | Provide a PowerShell script to run the example and capture retry logs.

using System;
using System.Data;
using System.IO;
using System.Threading;
using Aspose.Cells;

// Demonstrates how to add configurable retry logic around WorkbookDesigner.SetDataSource using a helper method, enabling resilient smart‑marker processing when temporary database connectivity failures occur.
static class Config
{
    // Simple configuration holder for retry attempts.
    public static int MaxRetryTimes { get; set; } = 3;
}

class SmartMarkerRetryExample
{
    // Executes an action with retry logic based on Config.MaxRetryTimes.
    private static void ExecuteWithRetry(Action action)
    {
        int maxRetries = Config.MaxRetryTimes > 0 ? Config.MaxRetryTimes : 3;
        int attempt = 0;

        while (true)
        {
            try
            {
                action();
                break; // Success, exit loop
            }
            catch (Exception ex)
            {
                attempt++;
                if (attempt > maxRetries)
                {
                    Console.WriteLine($"Maximum retry attempts ({maxRetries}) reached. Rethrowing exception.");
                    throw;
                }

                Console.WriteLine($"Transient error encountered (Attempt {attempt}/{maxRetries}): {ex.Message}");
                // Optional: introduce a delay before retrying.
                Thread.Sleep(1000 * attempt);
            }
        }
    }

    // Creates a sample DataTable to be used as a data source for smart markers.
    private static DataTable GetSampleData()
    {
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Department", typeof(string));

        table.Rows.Add(1, "Alice", "Finance");
        table.Rows.Add(2, "Bob", "HR");
        table.Rows.Add(3, "Charlie", "IT");

        return table;
    }

    static void Main()
    {
        try
        {
            // Configure maximum retry attempts (can be set elsewhere in the application).
            Config.MaxRetryTimes = 3;

            // Verify the template file exists before loading.
            string templatePath = "TemplateWithSmartMarkers.xlsx";
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the Excel template that contains smart markers.
            Workbook workbook = new Workbook(templatePath);

            // Initialize the WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Wrap the SetDataSource call with retry logic.
            ExecuteWithRetry(() =>
            {
                // Use a DataTable as the data source for smart markers.
                DataTable data = GetSampleData();
                designer.SetDataSource(data);
            });

            // Process the smart markers after the data source has been successfully set.
            designer.Process();

            // Save the populated workbook.
            string outputPath = "OutputWithSmartMarkers.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
