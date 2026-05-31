using System;
using System.Data;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsRetryDemo
{
    // Simple configuration holder.
    internal static class Config
    {
        // Maximum number of retry attempts for transient errors.
        public static int MaxRetryTimes { get; set; } = 3;
    }

    class Program
    {
        // Wrapper that retries SetDataSource when a transient error occurs.
        static void SetDataSourceWithRetry(WorkbookDesigner designer, string name, object dataSource)
        {
            int maxRetries = Config.MaxRetryTimes > 0 ? Config.MaxRetryTimes : 3;
            int attempt = 0;

            while (true)
            {
                try
                {
                    designer.SetDataSource(name, dataSource);
                    break; // Success
                }
                catch (Exception ex) when (IsTransient(ex))
                {
                    attempt++;
                    if (attempt > maxRetries) throw;

                    int delayMs = (int)Math.Pow(2, attempt) * 100;
                    Console.WriteLine($"Transient error on attempt {attempt}. Retrying in {delayMs} ms. Error: {ex.Message}");
                    Thread.Sleep(delayMs);
                }
            }
        }

        // Very basic check for transient errors (placeholder – always returns false in this demo).
        static bool IsTransient(Exception ex)
        {
            // Extend this method with real transient‑error detection as needed.
            return false;
        }

        static void Main()
        {
            try
            {
                const string templatePath = "TemplateWithSmartMarkers.xlsx";
                const string resultPath = "ResultWithSmartMarkers.xlsx";

                // Verify template file exists to avoid FileNotFoundException.
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook that contains smart markers.
                Workbook workbook = new Workbook(templatePath);
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Create a simple DataTable as a data source (replace with real data as needed).
                DataTable table = new DataTable("MyData");
                table.Columns.Add("Id", typeof(int));
                table.Columns.Add("Name", typeof(string));
                table.Rows.Add(1, "Alice");
                table.Rows.Add(2, "Bob");
                table.Rows.Add(3, "Charlie");

                // Bind the DataTable to the designer with retry logic.
                SetDataSourceWithRetry(designer, "MyData", table);

                // Process smart markers after the data source is successfully bound.
                designer.Process();

                // Save the populated workbook.
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved to {resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}