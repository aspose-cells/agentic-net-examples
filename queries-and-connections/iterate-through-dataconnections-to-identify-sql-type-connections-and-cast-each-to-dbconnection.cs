// Title: How to enumerate workbook DataConnections and list SQL DBConnection details using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loops through a Workbook's DataConnections, detects DBConnection objects with OLE DB or ODBC source types, and writes their Name, ConnectionString, and Command to the console. | Create a reusable method in Aspose.Cells that filters external connections for SQL‑type sources and returns their key properties as a collection.
// Common Searches: aspnet c# iterate Excel DataConnections to find ODBC SQL connections with Aspose.Cells | how to get DBConnection properties from a workbook using Aspose.Cells | filter external connections for SQL databases in an Excel file with Aspose.Cells .NET | list OLE DB and ODBC data connections in an Excel workbook using Aspose.Cells
// Tags: enumerate Aspose.Cells DataConnections | filter DBConnection by source type | display SQL connection properties C# | Aspose.Cells external connection inspection

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, accesses its DataConnections collection, iterates over each ExternalConnection, identifies DBConnection objects, filters for OLE DB or ODBC source types (common for SQL databases), prints relevant connection details to the console, and saves the workbook.
    class IdentifySqlConnections
    {
        static void Main()
        {
            // Load an existing workbook that may contain external data connections.
            // Replace "input.xlsx" with the path to your workbook.
            Workbook workbook = new Workbook("input.xlsx");

            // Get the collection of external connections from the workbook.
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Iterate through all connections.
            for (int i = 0; i < connections.Count; i++)
            {
                ExternalConnection conn = connections[i];

                // Check if the connection is a DBConnection (ODBC or OLE DB).
                if (conn is DBConnection dbConn)
                {
                    // Optional: further filter for SQL‑type sources.
                    // OLEDBBasedSource (value 5) and ODBCBasedSource (value 1) are typical for SQL databases.
                    bool isSqlSource = dbConn.SourceType == ConnectionDataSourceType.OLEDBBasedSource ||
                                       dbConn.SourceType == ConnectionDataSourceType.ODBCBasedSource;

                    if (isSqlSource)
                    {
                        Console.WriteLine($"SQL Connection found at index {i}:");
                        Console.WriteLine($"  Name               : {dbConn.Name}");
                        Console.WriteLine($"  ClassType          : {dbConn.ClassType}");
                        Console.WriteLine($"  SourceType         : {dbConn.SourceType}");
                        Console.WriteLine($"  ConnectionString   : {dbConn.ConnectionString}");
                        Console.WriteLine($"  CommandType        : {dbConn.CommandType}");
                        Console.WriteLine($"  Command            : {dbConn.Command}");
                    }
                }
            }

            // Save the workbook (even if unchanged) to demonstrate a complete flow.
            // Replace "output.xlsx" with the desired output path.
            workbook.Save("output.xlsx");
        }
    }
}
