// Title: C# example: add configurable retry mechanism for Aspose.Cells smart marker population from a network data source
// AI Prompts: Wrap the GetNetworkData call in a retry loop that retries up to a configurable maxAttempts with a delayBetweenAttempts, catching only transient WebException, then invoke WorkbookDesigner.Process(). | Implement an IsTransient helper that evaluates exceptions (e.g., WebException, TimeoutException) and integrate exponential back‑off into the retry logic for smart marker population. | Refactor the synchronous retry code to an async version using async/await, Task.Delay, and a CancellationToken while preserving WorkbookDesigner usage.
// Common Searches: c# Aspose.Cells retry smart markers after network failure | how to implement retry logic for WorkbookDesigner data source in C# | configurable max attempts and delay for smart marker population Aspose.Cells | example of exponential backoff with Aspose.Cells smart markers | handling transient WebException when populating smart markers in C#
// Tags: smart marker retry Aspose.Cells | WorkbookDesigner transient network handling | configurable retry loop C# Aspose.Cells | exponential backoff smart markers | network data source exception filtering Aspose.Cells

using System;
using System.Data;
using System.IO;
using System.Net;
using System.Threading;
using Aspose.Cells;

// Demonstrates loading a template workbook, retrieving data from a simulated network source, applying a configurable retry loop for transient WebException, processing smart markers via WorkbookDesigner, and saving the populated workbook.
class SmartMarkerRetryExample
{
    static void Main()
    {
        // Paths to the template and the output workbook
        string templatePath = "template.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the template file exists to avoid FileNotFoundException
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        // Load the workbook (lifecycle rule)
        Workbook workbook = new Workbook(templatePath);

        // Retry configuration
        int maxRetries = 3;               // maximum number of attempts
        int delayMilliseconds = 2000;    // wait time between attempts

        int attempt = 0;
        bool succeeded = false;

        while (attempt < maxRetries && !succeeded)
        {
            try
            {
                // Obtain data from a network‑based source (may throw transient errors)
                DataSet dataSource = GetNetworkData();

                // Populate smart markers using WorkbookDesigner (correct Aspose.Cells API)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSource);
                designer.Process();

                succeeded = true; // success, exit loop
            }
            catch (Exception ex) when (IsTransient(ex))
            {
                attempt++;
                if (attempt >= maxRetries)
                {
                    Console.WriteLine($"Operation failed after {attempt} attempts: {ex.Message}");
                    throw; // re‑throw after exhausting retries
                }

                Console.WriteLine($"Transient error encountered (attempt {attempt}/{maxRetries}). Retrying in {delayMilliseconds} ms...");
                Thread.Sleep(delayMilliseconds);
            }
        }

        // Save the populated workbook (lifecycle rule) with safety handling
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }

    // Simulated network call that returns data for smart markers.
    // Replace with actual HTTP/DB call as needed.
    private static DataSet GetNetworkData()
    {
        // Randomly simulate a transient network failure
        Random rnd = new Random();
        if (rnd.NextDouble() < 0.5)
        {
            throw new WebException("Simulated transient network failure.");
        }

        // Create a DataSet containing a DataTable matching smart marker fields
        DataSet ds = new DataSet();
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Rows.Add("John Doe", 30);
        table.Rows.Add("Jane Smith", 28);
        ds.Tables.Add(table);
        return ds;
    }

    // Determines whether an exception is transient and should be retried.
    private static bool IsTransient(Exception ex)
    {
        // For this example, treat WebException as transient.
        // Extend this method to cover other transient scenarios (e.g., timeout).
        return ex is WebException;
    }
}
