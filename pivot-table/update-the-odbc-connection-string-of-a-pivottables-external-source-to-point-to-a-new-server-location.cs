// Title: C# – Update PivotTable ODBC Connection String to a New Server with Aspose.Cells
// Description: Loads an Excel workbook, iterates through each worksheet and PivotTable, finds external ODBC connections, replaces the old server segment in the ConnectionString (and DBConnection.Command when present), and saves the workbook with the revised data source.
// Keywords: Aspose.Cells | C# | PivotTable | ODBC connection string | external data source | DBConnection | update server name | Excel automation | change data source programmatically
// Common Searches: Aspose.Cells change pivot table ODBC server | C# update external connection string in Excel workbook | replace old server in PivotTable data source | modify DBConnection command text Aspose | programmatically migrate pivot table data source
// Developer Intent: Replace the old server identifier in every PivotTable ODBC connection string (and related command text) within a workbook.
// Use Cases: Migrate a workbook’s pivot tables to a new database server after an upgrade. | Batch‑process multiple Excel files to correct outdated ODBC connection strings. | Ensure both ConnectionString and Command properties reflect the new server for ODBC/OLE DB sources.
// AI Prompts: Write C# code that scans all PivotTables in a workbook and updates the server part of their ODBC ConnectionString using Aspose.Cells. | Create a reusable method that accepts oldServerPart and newServerPart and updates ConnectionString and DBConnection.Command for every external connection in a workbook. | Explain how to validate that updated PivotTable connections are applied before saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

// Loads an Excel workbook, iterates through each worksheet and PivotTable, finds external ODBC connections, replaces the old server segment in the ConnectionString (and DBConnection.Command when present), and saves the workbook with the revised data source.
class UpdatePivotTableConnection
{
    static void Main()
    {
        // Load the workbook that contains the pivot table with an external ODBC source
        Workbook workbook = new Workbook("input.xlsx");

        // Define the part of the connection string that identifies the old server
        // and the replacement that points to the new server location
        string oldServerPart = "Server=OldServer;";
        string newServerPart = "Server=NewServer;";

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all pivot tables on the current worksheet
            foreach (PivotTable pivot in sheet.PivotTables)
            {
                // Retrieve the external connections used by this pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                // Update each connection's ConnectionString if it references the old server
                foreach (ExternalConnection conn in connections)
                {
                    if (!string.IsNullOrEmpty(conn.ConnectionString) &&
                        conn.ConnectionString.Contains(oldServerPart))
                    {
                        conn.ConnectionString = conn.ConnectionString.Replace(oldServerPart, newServerPart);
                    }

                    // If the connection is a DBConnection (ODBC/OLE DB), optionally update its Command text
                    if (conn is DBConnection dbConn && !string.IsNullOrEmpty(dbConn.Command))
                    {
                        dbConn.Command = dbConn.Command.Replace(oldServerPart, newServerPart);
                    }
                }
            }
        }

        // Save the workbook with the updated connection strings
        workbook.Save("output.xlsx");
    }
}
