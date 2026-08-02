// Title: Aspose.Cells for .NET – Turn Off QueryTable BackgroundRefresh (Synchronous Retrieval)
// Description: Loads a workbook, locates the first QueryTable, accesses its ExternalConnection, sets BackgroundRefresh = false to force synchronous data refresh, and saves the file. Demonstrates how to disable asynchronous background refresh in Aspose.Cells.
// Keywords: Aspose.Cells | QueryTable | BackgroundRefresh | synchronous data | C# example | ExternalConnection | disable background refresh | Excel data import
// Common Searches: Aspose.Cells disable query table background refresh | set BackgroundRefresh false C# | synchronous query table refresh Aspose.Cells | turn off background refresh Excel using Aspose | how to make query table refresh synchronously .NET
// Developer Intent: Disable the background refresh of a QueryTable so that data is fetched synchronously.
// Use Cases: Ensure query table data is fully refreshed before calculations or export. | Avoid race conditions in automated reporting pipelines that rely on up‑to‑date data. | Standardize connection settings across multiple workbooks for consistent data retrieval.
// AI Prompts: Write C# code with Aspose.Cells that disables BackgroundRefresh for all QueryTables in a workbook. | Explain how to confirm that BackgroundRefresh = false was saved in the workbook. | Provide error handling for QueryTables without an ExternalConnection when disabling background refresh.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Loads a workbook, locates the first QueryTable, accesses its ExternalConnection, sets BackgroundRefresh = false to force synchronous data refresh, and saves the file. Demonstrates how to disable asynchronous background refresh in Aspose.Cells.
class SetQueryTableBackgroundRefresh
{
    static void Main()
    {
        // Load an existing workbook that contains a query table
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the worksheet has at least one query table
        if (worksheet.QueryTables.Count > 0)
        {
            // Retrieve the first query table
            QueryTable queryTable = worksheet.QueryTables[0];

            // Obtain the external connection associated with the query table
            ExternalConnection externalConnection = queryTable.ExternalConnection;

            if (externalConnection != null)
            {
                // Set BackgroundRefresh to false to enforce synchronous data retrieval
                externalConnection.BackgroundRefresh = false;
                Console.WriteLine("BackgroundRefresh set to false for the query table.");
            }
            else
            {
                Console.WriteLine("The query table does not have an associated external connection.");
            }
        }
        else
        {
            Console.WriteLine("No query tables found in the worksheet.");
        }

        // Save the workbook with the updated connection settings
        workbook.Save("output.xlsx");
    }
}
