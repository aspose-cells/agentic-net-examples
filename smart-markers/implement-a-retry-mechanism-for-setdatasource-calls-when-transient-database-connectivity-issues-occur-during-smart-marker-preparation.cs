// Title: Implement retry logic for WorkbookDesigner.SetDataSource when handling transient database errors in Aspose.Cells smart markers (C#)
// Common Searches: C# retry pattern for Aspose.Cells WorkbookDesigner SetDataSource transient failures | How to handle database connectivity issues while setting smart marker data source in Aspose.Cells | Aspose.Cells smart markers retry on SetDataSource exception example | Implement exponential backoff for WorkbookDesigner SetDataSource in C# | Retrying SetDataSource call in Aspose.Cells when database query fails
// Tags: WorkbookDesigner SetDataSource retry | Aspose.Cells smart marker data source handling | C# transient database retry | Excel template processing with retry | smart markers error handling C#

using System;
using System.Data;
using System.IO;
using System.Threading;
using Aspose.Cells;

// The example loads an Excel template, obtains a DataTable, and uses a custom SetDataSourceWithRetry method that attempts to set the smart‑marker data source via WorkbookDesigner up to three times with a 2‑second pause between attempts. On success it processes the smart markers, recalculates formulas, and saves the result workbook.
class SmartMarkerProcessor
{
    // Maximum number of retry attempts
    private const int MaxRetryAttempts = 3;
    // Delay between retries (in milliseconds)
    private const int RetryDelay = 2000;

    static void Main()
    {
        try
        {
            // Load the template workbook (lifecycle rule)
            Workbook workbook;
            const string templatePath = "Template.xlsx";

            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                // If the template is missing, create a new workbook as a fallback
                Console.WriteLine($"Template file '{templatePath}' not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Prepare the data source (example DataTable)
            DataTable dt = GetDataFromDatabase();

            // Set the data source with retry logic and process smart markers
            SetDataSourceWithRetry(workbook, "MyDataSource", dt);

            // Recalculate formulas after processing smart markers
            workbook.CalculateFormula();

            // Save the result workbook (lifecycle rule)
            const string resultPath = "Result.xlsx";
            try
            {
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Global exception handling
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Retrieves data from a database (placeholder for actual DB call)
    private static DataTable GetDataFromDatabase()
    {
        // In a real scenario, replace this with actual DB access code.
        DataTable table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add(1, "Alice");
        table.Rows.Add(2, "Bob");
        return table;
    }

    // Sets a smart marker data source with retry on transient connectivity issues
    private static void SetDataSourceWithRetry(Workbook workbook, string dataSourceName, DataTable data)
    {
        int attempt = 0;
        while (true)
        {
            try
            {
                // Use WorkbookDesigner to work with smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSourceName, data);
                designer.Process(); // Process smart markers
                break; // Success
            }
            catch (Exception ex)
            {
                attempt++;
                if (attempt >= MaxRetryAttempts)
                {
                    // Re‑throw after exceeding max attempts
                    throw new InvalidOperationException($"Failed to set data source after {MaxRetryAttempts} attempts.", ex);
                }

                // Log the error and wait before retrying
                Console.WriteLine($"Error setting data source (attempt {attempt}): {ex.Message}");
                Thread.Sleep(RetryDelay);
            }
        }
    }
}
