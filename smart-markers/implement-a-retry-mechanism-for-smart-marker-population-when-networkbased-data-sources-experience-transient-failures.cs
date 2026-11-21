// Title: Add retry mechanism for smart marker population from a network data source using Aspose.Cells in C#
// Common Searches: c# Aspose.Cells retry smart marker population when web service fails | how to implement transient error handling for WorkbookDesigner SetDataSource | retry pattern for fetching network data for smart markers in Excel | example of exponential backoff with Aspose.Cells smart markers | handling intermittent network failures while populating Excel templates using Aspose.Cells
// Tags: smart marker retry Aspose.Cells | WorkbookDesigner SetDataSource retry pattern | transient network exception handling Excel generation | C# retry mechanism for data-driven workbook | Aspose.Cells resilience to network failures

using System;
using System.Data;
using System.IO;
using System.Net;
using System.Threading;
using Aspose.Cells;

// The example loads a template workbook, retrieves data from a simulated network source inside a retry loop (up to three attempts with a 2‑second pause), populates smart markers via WorkbookDesigner, and saves the resulting Excel file, rethrowing if the retry limit is exceeded.
class SmartMarkerRetryExample
{
    static void Main()
    {
        try
        {
            // Verify that the template file exists before loading
            const string templatePath = "Template.xlsx";
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Retry configuration
            const int maxRetries = 3;               // maximum number of attempts
            const int delayMilliseconds = 2000;     // wait time between attempts
            int attempt = 0;
            bool succeeded = false;

            while (attempt < maxRetries && !succeeded)
            {
                try
                {
                    // Retrieve data from a network‑based source (may throw transient errors)
                    DataTable data = GetNetworkData();

                    // Populate smart markers using the retrieved data
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);
                    designer.SetDataSource(data);
                    designer.Process();

                    succeeded = true; // population succeeded, exit loop
                }
                catch (Exception ex) when (IsTransient(ex))
                {
                    // Transient failure – increment attempt counter and wait before retrying
                    attempt++;
                    if (attempt >= maxRetries)
                    {
                        // Re‑throw after exceeding retry limit
                        throw;
                    }
                    Thread.Sleep(delayMilliseconds);
                }
            }

            // Save the populated workbook
            const string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simulated method that fetches data from a network source.
    // In a real scenario this could be a call to a web service, REST API, etc.
    static DataTable GetNetworkData()
    {
        // Example: randomly throw a transient network exception
        Random rnd = new Random();
        if (rnd.NextDouble() < 0.5)
        {
            throw new WebException("Transient network failure");
        }

        // Build a simple DataTable as the data source for smart markers
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Salary", typeof(double));

        table.Rows.Add("John Doe", 50000);
        table.Rows.Add("Jane Smith", 60000);
        return table;
    }

    // Determines whether an exception is considered transient.
    // Extend this method to include other transient exception types as needed.
    static bool IsTransient(Exception ex)
    {
        return ex is WebException;
    }
}
