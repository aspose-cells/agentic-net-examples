// Title: Enumerate Excel DataConnections and extract SQL DBConnection details with Aspose.Cells for .NET
// Description: Load a workbook, access its DataConnections collection, iterate over each ExternalConnection, identify SQL‑type DBConnection objects using pattern matching, cast them, and read properties such as Name, ConnectionString and CommandType before saving the file.
// Keywords: Aspose.Cells | DataConnections | DBConnection | SQL external connection | C# .NET | iterate workbook connections | cast ExternalConnection to DBConnection | Excel data source properties | external data connection handling
// Common Searches: How to loop through DataConnections in Aspose.Cells | Filter SQL DBConnection objects in an Excel workbook using C# | Cast ExternalConnection to DBConnection with Aspose.Cells | Retrieve connection string from Excel data connections .NET | Aspose.Cells enumerate external connections example
// Developer Intent: The developer needs to enumerate all external data connections in an Excel workbook, detect which ones are SQL (DBConnection) type, and access their specific properties for further processing.
// Use Cases: Audit all SQL data sources in a workbook by logging their names and connection strings. | Modify the CommandType of each DBConnection before refreshing data to ensure correct query execution. | Generate a summary report of DBConnection details for compliance or migration tasks.
// AI Prompts: Write C# code that updates the ConnectionString of every DBConnection in a workbook's DataConnections collection using Aspose.Cells. | Create a method that returns a List<DBConnection> containing all SQL‑type connections from a given Workbook. | Explain how to handle authentication (username/password) for DBConnection objects while iterating through DataConnections in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Load a workbook, access its DataConnections collection, iterate over each ExternalConnection, identify SQL‑type DBConnection objects using pattern matching, cast them, and read properties such as Name, ConnectionString and CommandType before saving the file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of external data connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Iterate through all connections in the workbook
        for (int i = 0; i < connections.Count; i++)
        {
            ExternalConnection conn = connections[i];

            // Identify connections that are of SQL type (DBConnection)
            if (conn is DBConnection dbConn)
            {
                // Cast succeeded – you can now work with DBConnection members
                Console.WriteLine($"DBConnection #{i + 1}");
                Console.WriteLine($"  Name: {dbConn.Name}");
                Console.WriteLine($"  ConnectionString: {dbConn.ConnectionString}");
                Console.WriteLine($"  CommandType: {dbConn.CommandType}");
                // Additional DBConnection handling can be added here
            }
        }

        // Save the workbook (even if unchanged) to demonstrate the complete flow
        workbook.Save("output.xlsx");
    }
}
