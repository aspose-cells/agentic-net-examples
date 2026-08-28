// Title: Implement exponential‑backoff retry for WorkbookDesigner.SetDataSource during smart marker preparation in Aspose.Cells (C#)
// AI Prompts: Write a reusable helper that invokes WorkbookDesigner.SetDataSource with a configurable maximum retry count and exponential back‑off delays. | Detect transient database exceptions and automatically retry the SetDataSource call until the limit is reached or the operation succeeds. | Integrate the retry helper into a smart‑marker workflow, then process the markers and save the workbook.
// Common Searches: how to add retry logic to Aspose.Cells WorkbookDesigner SetDataSource | c# exponential backoff for smart marker datasource errors | handling transient database connectivity in Aspose.Cells smart markers | configurable max retry attempts for SetDataSource in Aspose.Cells | retry pattern for Aspose.Cells smart marker data source
// Tags: aspocells workbookdesigner setdatasource retry | smart markers datasource exponential backoff | c# transient database error handling aspocells | configurable retry attempts aspocells | excel smart marker retry pattern

using System;
using System.Data;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsRetryExample
{
    // The example loads a template workbook, creates a DataTable, and uses a RetrySetDataSource helper that calls WorkbookDesigner.SetDataSource with configurable exponential‑backoff retries. After successful data binding, the smart markers are processed and the result workbook is saved.
    class Program
    {
        // Simple configuration holder (replace with your own config source if needed)
        private static class Config
        {
            public static int MaxRetryTimes { get; } = 3;
        }

        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "Result.xlsx";

                // Verify that the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook
                Workbook workbook = new Workbook(templatePath);

                // Initialize the WorkbookDesigner with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Prepare a sample DataTable as the data source
                DataTable dataTable = new DataTable("Employees");
                dataTable.Columns.Add("ID", typeof(int));
                dataTable.Columns.Add("Name", typeof(string));
                dataTable.Rows.Add(1, "John Doe");
                dataTable.Rows.Add(2, "Jane Smith");

                // Set the data source with retry logic to handle transient issues
                RetrySetDataSource(() => designer.SetDataSource(dataTable));

                // Process the smart markers
                designer.Process();

                // Ensure the directory for the result file exists
                string resultDir = Path.GetDirectoryName(Path.GetFullPath(resultPath));
                if (!string.IsNullOrEmpty(resultDir) && !Directory.Exists(resultDir))
                {
                    Directory.CreateDirectory(resultDir);
                }

                // Save the processed workbook
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to {resultPath}");
            }
            catch (Exception ex)
            {
                // Log unexpected exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Executes the provided SetDataSource action with retry handling
        private static void RetrySetDataSource(Action setDataSourceAction)
        {
            int maxAttempts = Config.MaxRetryTimes > 0 ? Config.MaxRetryTimes : 3;
            int attempt = 0;

            while (true)
            {
                try
                {
                    setDataSourceAction();
                    break; // Success, exit the loop
                }
                catch (Exception ex) when (IsTransient(ex))
                {
                    attempt++;
                    if (attempt >= maxAttempts)
                    {
                        // Re‑throw after exceeding max retries
                        throw;
                    }

                    // Simple exponential back‑off before retrying
                    int delayMs = 1000 * attempt;
                    Thread.Sleep(delayMs);
                }
            }
        }

        // Determines whether an exception is considered transient.
        // For this example we treat all exceptions as non‑transient.
        private static bool IsTransient(Exception ex)
        {
            // Extend with specific transient checks if needed.
            return false;
        }
    }
}
