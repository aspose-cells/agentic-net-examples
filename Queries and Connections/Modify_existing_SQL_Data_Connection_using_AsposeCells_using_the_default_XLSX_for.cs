using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class ModifySqlConnection
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of external data connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Find the SQL (DB) connection and modify its properties
        foreach (ExternalConnection conn in connections)
        {
            if (conn is DBConnection dbConn)
            {
                // Show the current connection string
                Console.WriteLine("Original ConnectionString: " + dbConn.ConnectionString);

                // Update the connection string (example for a SQL Server)
                dbConn.ConnectionString = "Provider=SQLOLEDB;Data Source=MyServer;Initial Catalog=MyDatabase;Integrated Security=SSPI;";

                // Optionally update the command/query
                dbConn.Command = "SELECT * FROM dbo.MyTable";

                // Confirm the update
                Console.WriteLine("Updated ConnectionString: " + dbConn.ConnectionString);
            }
        }

        // Save the workbook with the modified connection
        workbook.Save("output.xlsx");
    }
}