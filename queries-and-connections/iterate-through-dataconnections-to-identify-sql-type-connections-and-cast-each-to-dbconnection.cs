using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class IdentifySqlConnections
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of external data connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Iterate through all connections in the workbook
        for (int i = 0; i < connections.Count; i++)
        {
            ExternalConnection conn = connections[i];

            // Check whether the connection is a DBConnection (ODBC/OLE DB)
            if (conn is DBConnection dbConn)
            {
                // DBConnection.CommandType == SqlStatement indicates a SQL‑based connection
                if (dbConn.CommandType == OLEDBCommandType.SqlStatement)
                {
                    Console.WriteLine($"Connection #{i} ('{dbConn.Name}') is a SQL type connection.");
                    Console.WriteLine($"  ConnectionString: {dbConn.ConnectionString}");
                }
                else
                {
                    Console.WriteLine($"Connection #{i} ('{dbConn.Name}') is a DBConnection but not a SQL statement (CommandType = {dbConn.CommandType}).");
                }
            }
        }

        // Save the workbook (if any changes were made)
        workbook.Save("output.xlsx");
    }
}