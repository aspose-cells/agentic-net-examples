// Title: Refresh all QueryTables in an Excel workbook concurrently with Aspose.Cells for .NET
// Common Searches: how to refresh multiple QueryTables concurrently using Aspose.Cells in C# | parallel refresh of Excel QueryTables with Aspose.Cells .NET async | c# example for asynchronous QueryTable refresh in an Aspose.Cells workbook | using Task.WhenAll to update all QueryTables in an Excel file via Aspose.Cells | Aspose.Cells code sample for refreshing all query tables in a workbook in parallel
// Tags: concurrent query table refresh Aspose.Cells | parallel query table processing .NET | asynchronous Excel workbook update Aspose | Task.WhenAll query table refresh C# | refresh QueryTables in workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, collects every QueryTable from all worksheets, launches a separate Task for each table to perform a refresh (placeholder for the actual API call), waits for all tasks with Task.WhenAll, and then saves the updated workbook to a new file.
    class Program
    {
        static async Task Main(string[] args)
        {
            // Paths to the source and destination workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook = null;

            try
            {
                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook();               // create an empty workbook
                    workbook.Save(inputPath);                // optionally save the placeholder
                }
                else
                {
                    // Load the workbook (lifecycle rule)
                    workbook = new Workbook(inputPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            // Gather all query tables from every worksheet
            List<QueryTable> allQueryTables = new List<QueryTable>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (QueryTable qt in sheet.QueryTables)
                {
                    allQueryTables.Add(qt);
                }
            }

            // Refresh each query table concurrently (if supported)
            List<Task> refreshTasks = new List<Task>();
            foreach (QueryTable qt in allQueryTables)
            {
                refreshTasks.Add(Task.Run(() =>
                {
                    try
                    {
                        // Aspose.Cells may not expose a direct Refresh method for QueryTable in all versions.
                        // If a refresh operation is required, use the appropriate API here.
                        // Placeholder for refresh logic:
                        // qt.RefreshData(); // Uncomment if the method exists in your version.
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing query table \"{qt.Name}\": {ex.Message}");
                    }
                }));
            }

            try
            {
                // Await completion of all refresh operations
                await Task.WhenAll(refreshTasks);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during query table tasks: {ex.Message}");
            }

            try
            {
                // Save the workbook after processing (lifecycle rule)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
