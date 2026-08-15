// Title: Change PivotTable External DB Connection to a New Server with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, scans every worksheet for PivotTables, extracts their DBConnection objects, replaces the old server identifier in the connection string, command text, and source file path, refreshes each PivotTable, and saves the updated workbook.
// Keywords: Aspose.Cells pivot table external connection | C# update DBConnection server | modify Excel pivot data source | refresh pivot tables after connection change | replace server name in ODBC connection string
// Common Searches: how to update pivot table server name using Aspose.Cells | C# change external database connection for Excel pivot tables | Aspose.Cells refresh pivot after DBConnection edit | replace old server with new in Excel pivot data source
// Developer Intent: Redirect all PivotTables in a workbook to a new database server and refresh them programmatically.
// Use Cases: Migrate reporting workbooks to a renamed or relocated SQL server without rebuilding PivotTables. | Automate connection updates across multiple Excel files in a CI/CD pipeline. | Ensure PivotTables display data from the new server immediately after deployment.
// AI Prompts: Write C# code with Aspose.Cells that swaps an old server name for a new one in every DBConnection of a workbook's PivotTables and then refreshes them. | Explain how to validate that each PivotTable uses the updated connection string before saving the file. | Suggest robust error‑handling for missing or malformed server tokens when updating PivotTable connections with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

// Loads an Excel workbook, scans every worksheet for PivotTables, extracts their DBConnection objects, replaces the old server identifier in the connection string, command text, and source file path, refreshes each PivotTable, and saves the updated workbook.
class ModifyPivotTableConnection
{
    static void Main()
    {
        // Load the workbook that contains the PivotTable with an external DB connection
        Workbook workbook = new Workbook("input.xlsx");

        // Define the new server name (or full connection string) you want the PivotTable to use
        string oldServer = "Server=OLD_SERVER;";
        string newServer = "Server=NEW_SERVER;";

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all pivot tables in the worksheet
            foreach (PivotTable pivot in sheet.PivotTables)
            {
                // Get all external connections used by this pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                foreach (ExternalConnection conn in connections)
                {
                    // We're interested only in database connections (ODBC/OLEDB)
                    if (conn is DBConnection dbConn)
                    {
                        // Update the connection string to point to the new server
                        // This simple replace assumes the server part is present in the string
                        if (!string.IsNullOrEmpty(dbConn.ConnectionString) && dbConn.ConnectionString.Contains(oldServer))
                        {
                            dbConn.ConnectionString = dbConn.ConnectionString.Replace(oldServer, newServer);
                        }

                        // Optionally, also update the command text if it contains server-specific paths
                        if (!string.IsNullOrEmpty(dbConn.Command))
                        {
                            dbConn.Command = dbConn.Command.Replace("OLD_SERVER", "NEW_SERVER");
                        }

                        // If the connection uses a source file path, update that as well
                        if (!string.IsNullOrEmpty(dbConn.SourceFile))
                        {
                            dbConn.SourceFile = dbConn.SourceFile.Replace("OLD_SERVER", "NEW_SERVER");
                        }
                    }
                }

                // Refresh the pivot table to apply the new connection settings
                pivot.RefreshData();
                pivot.CalculateData();
            }
        }

        // Refresh all pivot tables in the workbook (optional but ensures consistency)
        workbook.Worksheets.RefreshPivotTables();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
