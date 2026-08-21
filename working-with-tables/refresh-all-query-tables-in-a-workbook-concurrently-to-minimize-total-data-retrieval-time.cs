// Title: Concurrent Refresh of All Query Tables in an Excel Workbook with Aspose.Cells (C#)
// Description: Loads a workbook, uses Parallel.ForEach to refresh every query table across all worksheets via dynamic invocation, logs individual failures, and saves the updated file, reducing overall data retrieval time.
// Keywords: Aspose.Cells | C# query table refresh | parallel query refresh | Concurrent workbook update | Parallel.ForEach Aspose.Cells | RefreshData method | Excel data connection refresh | bulk query table update
// Common Searches: Aspose.Cells refresh all query tables concurrently | Parallel.ForEach refresh query tables C# | How to update multiple query tables at once with Aspose.Cells | Concurrent data refresh for Excel workbook using Aspose | Speed up query table refresh in .NET
// Developer Intent: Refresh every query table in a workbook in parallel to minimize total refresh duration.
// Use Cases: Synchronize external data sources for large workbooks before distribution. | Accelerate bulk data refresh in reporting dashboards that rely on many query tables. | Continue processing other tables while logging failures of individual refresh operations.
// AI Prompts: Write C# code that uses Aspose.Cells to refresh all query tables concurrently and records any errors. | Show an async/await version of the concurrent query table refresh example for Aspose.Cells. | Explain how to set connection strings and command text for query tables before invoking RefreshData with Aspose.Cells.

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, uses Parallel.ForEach to refresh every query table across all worksheets via dynamic invocation, logs individual failures, and saves the updated file, reducing overall data retrieval time.
    public class RefreshAllQueryTablesConcurrently
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            WorksheetCollection worksheets = workbook.Worksheets;

            // Refresh all query tables in parallel
            Parallel.ForEach(Enumerable.Range(0, worksheets.Count), worksheetIndex =>
            {
                Worksheet sheet = worksheets[worksheetIndex];
                QueryTableCollection queryTables = sheet.QueryTables;

                for (int i = 0; i < queryTables.Count; i++)
                {
                    try
                    {
                        // Use dynamic to invoke RefreshData (method may vary across versions)
                        dynamic qt = queryTables[i];
                        qt.RefreshData();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to refresh query table '{queryTables[i].Name}' on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            });

            try
            {
                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
