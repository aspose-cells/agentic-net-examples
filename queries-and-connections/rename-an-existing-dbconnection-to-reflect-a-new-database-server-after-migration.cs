using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class RenameDbConnection
{
    static void Main()
    {
        // Load the workbook that contains the DBConnection
        Workbook workbook = new Workbook("input.xlsx");

        // Define old and new server identifiers
        string oldServer = "OldServer";
        string newServer = "NewServer";

        // Iterate through all external connections in the workbook
        foreach (ExternalConnection conn in workbook.DataConnections)
        {
            // Process only DBConnection objects
            if (conn is DBConnection dbConn)
            {
                // Rename the connection if its name contains the old server identifier
                if (!string.IsNullOrEmpty(dbConn.Name) && dbConn.Name.Contains(oldServer))
                {
                    dbConn.Name = dbConn.Name.Replace(oldServer, newServer);
                }

                // Update the connection string to point to the new server
                if (!string.IsNullOrEmpty(dbConn.ConnectionString))
                {
                    dbConn.ConnectionString = dbConn.ConnectionString
                        .Replace($"Data Source={oldServer}", $"Data Source={newServer}")
                        .Replace($"Server={oldServer}", $"Server={newServer}");
                }
            }
        }

        // Save the workbook with the updated connection information
        workbook.Save("output.xlsx");
    }
}