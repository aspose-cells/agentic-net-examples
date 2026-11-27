// Title: How to set a 30‑second CommandTimeout on a DBConnection in Aspose.Cells for .NET via reflection
// Common Searches: aspnet set commandtimeout on dbconnection aspose.cells c# | c# aspose.cells external connection limit query execution time | using reflection to assign CommandTimeout property on Aspose.Cells DBConnection | aspose.cells dbconnection command timeout unavailable in older SDK versions | disable background refresh for Aspose.Cells DBConnection query execution
// Tags: Aspose.Cells DBConnection CommandTimeout | reflection set property C# Aspose.Cells | disable background refresh Aspose.Cells | external data connection timeout Aspose.Cells | configure DBConnection query timeout .NET

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsCommandTimeoutDemo
{
    // The example creates a new Workbook, adds an uninitialized DBConnection, configures its connection string, command type, and SQL command, then uses reflection to set a 30‑second CommandTimeout (handling SDKs where the property is missing), disables background refresh for synchronous execution, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output workbook
            const string outputPath = "DBConnection_CommandTimeoutDemo.xlsx";

            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the collection of external connections
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Add a new DBConnection to the collection (uninitialized object)
                DBConnection dbConnection = (DBConnection)connections[connections.Count];

                // ------------------------------------------------------------
                // Configure the DBConnection
                // ------------------------------------------------------------

                // 1. Set the connection string (example for OLE DB)
                dbConnection.ConnectionString = "Provider=SQLOLEDB;Data Source=YOUR_SERVER;Initial Catalog=YOUR_DATABASE;Integrated Security=SSPI;";

                // 2. Set the command type (SQL statement)
                dbConnection.CommandType = OLEDBCommandType.SqlStatement;

                // 3. Set the SQL command to be executed
                dbConnection.Command = "SELECT TOP 10 * FROM dbo.YourTable";

                // 4. Set the command timeout to 30 seconds via reflection (property may not exist in older SDKs)
                try
                {
                    var prop = typeof(DBConnection).GetProperty("CommandTimeout");
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(dbConnection, 30);
                    }
                    else
                    {
                        Console.WriteLine("Warning: CommandTimeout property is not available on DBConnection in this SDK version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error setting CommandTimeout via reflection: {ex.Message}");
                }

                // 5. Set background refresh to false to ensure synchronous execution
                dbConnection.BackgroundRefresh = false;

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("DBConnection configuration:");
                Console.WriteLine($"  ConnectionString : {dbConnection.ConnectionString}");
                Console.WriteLine($"  CommandType      : {dbConnection.CommandType}");
                Console.WriteLine($"  Command          : {dbConnection.Command}");
                Console.WriteLine($"  CommandTimeout   : 30 seconds (if supported)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while configuring DBConnection: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
