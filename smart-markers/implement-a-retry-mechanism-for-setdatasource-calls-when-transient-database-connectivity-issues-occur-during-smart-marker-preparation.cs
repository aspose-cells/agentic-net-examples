// Title: Implement exponential back‑off retry for WorkbookDesigner.SetDataSource with Aspose.Cells smart markers in C#
// Common Searches: C# how to retry WorkbookDesigner.SetDataSource on transient database errors | Aspose.Cells smart markers exponential backoff example for data source binding | Implement retry mechanism for SetDataSource during smart marker preparation in .NET | Resilient Excel template processing with transient error handling in C# | Best practice for handling transient connectivity when using Aspose.Cells smart markers
// Tags: Aspose.Cells WorkbookDesigner resilient data binding | smart markers retry delay strategy | transient connectivity handling in Excel automation | C# SetDataSource exception handling pattern | Excel template processing fallback data source

using System;
using System.Data;
using System.IO;
using System.Threading;
using Aspose.Cells;

// The sample loads an Excel template, creates a WorkbookDesigner, retrieves employee data into a DataTable, and sets the smart marker data source using a custom SetDataSourceWithRetry method. The method retries transient failures with exponential back‑off, ensuring reliable data binding before processing smart markers and saving the result workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the workbook template that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Create a WorkbookDesigner for processing smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Obtain data from the database (placeholder implementation)
            DataTable employeeTable = GetEmployeeData();

            // Set the data source for the smart marker with a retry mechanism
            SetDataSourceWithRetry(designer, "Employees", employeeTable, maxRetries: 3, initialDelayMs: 2000);

            // Process the smart markers after the data source has been set
            designer.Process();

            // Ensure the output directory exists
            string resultDir = Path.GetDirectoryName(Path.GetFullPath(resultPath));
            if (!string.IsNullOrEmpty(resultDir) && !Directory.Exists(resultDir))
                Directory.CreateDirectory(resultDir);

            // Save the resulting workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Retries SetDataSource when a transient exception occurs
    static void SetDataSourceWithRetry(WorkbookDesigner designer, string dataSourceName, object dataSource, int maxRetries, int initialDelayMs)
    {
        int attempt = 0;
        int delay = initialDelayMs;

        while (true)
        {
            try
            {
                // Attempt to set the data source for the smart marker
                designer.SetDataSource(dataSourceName, dataSource);
                break; // Success – exit the retry loop
            }
            catch (Exception ex) when (IsTransient(ex))
            {
                attempt++;

                if (attempt > maxRetries)
                {
                    // Exceeded the allowed number of retries; rethrow the exception
                    throw;
                }

                // Wait before the next retry (exponential back‑off)
                Thread.Sleep(delay);
                delay *= 2;
            }
        }
    }

    // Determines whether an exception is transient (placeholder logic)
    static bool IsTransient(Exception ex)
    {
        // In a real scenario, inspect the exception type/message (e.g., SqlException with specific error numbers)
        // For demonstration purposes, treat all exceptions as transient
        return true;
    }

    // Placeholder method that simulates retrieving data from a database
    static DataTable GetEmployeeData()
    {
        DataTable table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Rows.Add("John Doe", 30);
        table.Rows.Add("Jane Smith", 28);
        return table;
    }
}
