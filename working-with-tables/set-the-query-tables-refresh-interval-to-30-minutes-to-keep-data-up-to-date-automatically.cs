// Title: C# – Set Query Table Auto‑Refresh Interval to 30 Minutes with Aspose.Cells for .NET
// Description: Shows how to create a workbook, detect a query table on the first worksheet, access its ExternalConnection, set the automatic refresh interval to 30 minutes, and save the updated file.
// Keywords: Aspose.Cells | C# query table refresh | set refresh interval | ExternalConnection RefreshInterval | auto refresh query table | 30‑minute refresh | Aspose.Cells for .NET | Excel workbook refresh settings
// Common Searches: Aspose.Cells set query table refresh interval | C# set external connection refresh interval Aspose | how to auto refresh query table every 30 minutes | verify query tables before setting refresh interval Aspose.Cells | save workbook with query table refresh settings
// Developer Intent: Set the query table’s automatic refresh interval to 30 minutes.
// Use Cases: Automatically update data imported via a query table every 30 minutes. | Avoid runtime errors by confirming a query table exists before applying the setting. | Persist the auto‑refresh configuration when saving the workbook. | Integrate periodic data refresh into reporting or dashboard solutions.
// AI Prompts: Generate C# code using Aspose.Cells that checks for a query table, accesses its ExternalConnection, and sets the RefreshInterval to 30 minutes. | Provide an example that adds a new query table with an external connection and configures it to refresh every 30 minutes in Aspose.Cells. | Explain how to read back the RefreshInterval value to verify the change and how to modify it programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Shows how to create a workbook, detect a query table on the first worksheet, access its ExternalConnection, set the automatic refresh interval to 30 minutes, and save the updated file.
class SetQueryTableRefreshInterval
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Check if there is at least one query table in the worksheet
        if (worksheet.QueryTables.Count > 0)
        {
            // Get the first query table
            QueryTable queryTable = worksheet.QueryTables[0];

            // Retrieve the external connection associated with the query table
            ExternalConnection connection = queryTable.ExternalConnection;

            // Set the automatic refresh interval to 30 minutes
            connection.RefreshInternal = 30;

            Console.WriteLine("Refresh interval set to " + connection.RefreshInternal + " minutes.");
        }
        else
        {
            Console.WriteLine("No query tables found in the worksheet.");
        }

        // Save the workbook with the updated settings
        workbook.Save("QueryTableRefreshInterval.xlsx");
    }
}
