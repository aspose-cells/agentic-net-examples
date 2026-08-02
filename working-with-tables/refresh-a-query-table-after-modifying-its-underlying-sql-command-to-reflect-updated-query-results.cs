// Title: Refresh a Query Table After Changing Its SQL Command with Aspose.Cells for .NET
// Description: This example loads a workbook, retrieves the first query table, updates the linked ExternalConnection's SQL statement, enables formatting preservation and column‑width auto‑adjust, refreshes all external connections, and saves the result to a new file.
// Keywords: Aspose.Cells query table refresh | modify external connection SQL | C# Workbook.RefreshAll | preserve formatting after refresh | adjust column width .NET | update query table data programmatically
// Common Searches: change SQL command of a query table using Aspose.Cells | how to refresh query tables after editing the command in C# | keep cell formatting when refreshing external connections | auto‑fit columns after query table refresh .NET | Aspose.Cells refresh all connections example
// Developer Intent: Update a query table's SQL query and reload its data programmatically.
// Use Cases: Switch a query table to retrieve only active records, refresh, and keep the original styling. | Batch‑update multiple query tables' commands in one workbook and refresh them in a single call. | Save a refreshed workbook with columns automatically resized to accommodate new data.
// AI Prompts: Write C# code that changes a query table's SQL command, calls Workbook.RefreshAll, and saves the workbook while preserving formatting with Aspose.Cells. | Explain how ExternalConnection.Command interacts with Workbook.RefreshAll when updating query tables. | Suggest robust error‑handling for missing input files, absent query tables, or invalid SQL statements in an Aspose.Cells workflow.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // This example loads a workbook, retrieves the first query table, updates the linked ExternalConnection's SQL statement, enables formatting preservation and column‑width auto‑adjust, refreshes all external connections, and saves the result to a new file.
    public class RefreshQueryTableDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook that contains a query table.
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed).
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one query table in the worksheet.
                if (worksheet.QueryTables.Count == 0)
                {
                    Console.WriteLine("No query tables found in the worksheet.");
                    return;
                }

                // Get the first query table.
                QueryTable queryTable = worksheet.QueryTables[0];

                // Obtain the external connection linked to the query table.
                ExternalConnection connection = queryTable.ExternalConnection;

                // Modify the SQL command text to reflect the updated query.
                connection.Command = "SELECT * FROM dbo.YourTable WHERE Status = 'Active'";

                // Preserve formatting and adjust column width after refresh.
                queryTable.PreserveFormatting = true;
                queryTable.AdjustColumnWidth = true;

                // Refresh all external connections (including the query table) to apply the new command.
                workbook.RefreshAll();

                // Save the workbook with the refreshed data.
                workbook.Save(outputPath);
                Console.WriteLine($"Query table refreshed and workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
