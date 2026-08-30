// Title: Retrieve the Name of the first DBConnection from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file with Aspose.Cells, get the workbook's DataConnections collection, cast the first connection to DBConnection, and output its Name property. | Check whether the first external connection in a workbook is a DBConnection and, if so, print its Name value. | Iterate the workbook's external connections, identify the first DBConnection, and retrieve its Name attribute using C#.
// Common Searches: Aspose.Cells C# get name of first database connection in workbook | How to read DBConnection Name property from Excel file using Aspose.Cells .NET | Retrieve external data connection name from .xlsx with Aspose.Cells API
// Tags: read DBConnection Name Aspose.Cells .NET | cast first external connection to DBConnection C# | workbook DataConnections enumeration Aspose.Cells | extract database connection identifier from Excel file | manage external data connections Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The example loads an existing .xlsx workbook with Aspose.Cells, accesses its DataConnections collection, verifies that the first external connection is a DBConnection, reads the connection's Name property, prints it to the console, and optionally saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (provide the correct file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        if (connections.Count > 0)
        {
            // Retrieve the first external connection
            ExternalConnection firstConnection = connections[0];

            // Verify that the connection is a DBConnection
            if (firstConnection is DBConnection dbConnection)
            {
                // Read the Name property of the DBConnection
                string connectionName = dbConnection.Name;
                Console.WriteLine("First DBConnection Name: " + connectionName);
            }
            else
            {
                Console.WriteLine("The first connection is not a DBConnection. Type: " + firstConnection.GetType().Name);
            }
        }
        else
        {
            Console.WriteLine("No external connections found in the workbook.");
        }

        // Save the workbook (optional, adjust the path as needed)
        workbook.Save("output.xlsx");
    }
}
