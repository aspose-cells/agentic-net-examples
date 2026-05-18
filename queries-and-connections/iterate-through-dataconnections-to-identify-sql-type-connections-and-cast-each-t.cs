using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Load an existing workbook that may contain external connections.
        // Replace the path with the actual file you want to process.
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of external data connections.
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Iterate through each connection in the collection.
        for (int i = 0; i < connections.Count; i++)
        {
            ExternalConnection conn = connections[i];

            // Identify connections that are of type DBConnection (ODBC/OLE DB).
            if (conn is DBConnection dbConn)
            {
                // Further filter to SQL statements if needed.
                // OLEDBCommandType.SqlStatement indicates a SQL query.
                if (dbConn.CommandType == OLEDBCommandType.SqlStatement)
                {
                    Console.WriteLine($"SQL DBConnection found at index {i}:");
                    Console.WriteLine($"  Name             : {dbConn.Name}");
                    Console.WriteLine($"  ConnectionString : {dbConn.ConnectionString}");
                    Console.WriteLine($"  Command          : {dbConn.Command}");
                }
            }
        }

        // Save the workbook (even if no changes were made) to demonstrate a complete flow.
        workbook.Save("output.xlsx");
    }
}