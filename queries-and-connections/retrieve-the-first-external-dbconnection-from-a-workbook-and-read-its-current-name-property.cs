// Title: Get the First DBConnection Name from an Excel Workbook with Aspose.Cells for .NET
// Description: Loads a workbook, accesses its DataConnections collection, verifies the presence of at least one connection, casts the first item to DBConnection, reads the Name property, prints it to the console, and optionally saves the file.
// Keywords: Aspose.Cells DBConnection name | read external connection .NET | first data connection Excel | retrieve DBConnection Name property | Aspose.Cells external connections example
// Common Searches: Aspose.Cells get first DBConnection name | read external DBConnection property C# | how to access workbook data connections Aspose | retrieve connection name from Excel file using Aspose.Cells
// Developer Intent: Extract the Name value of the initial DBConnection defined in an Excel workbook.
// Use Cases: Log the connection identifier for troubleshooting before data import. | Validate that a workbook references the expected database by comparing the retrieved name. | Show the source connection name in a UI to inform end‑users about the linked database.
// AI Prompts: Generate C# code that iterates through all DBConnection objects in a workbook and outputs each Name. | Demonstrate how to rename a DBConnection after reading its current Name and then save the workbook. | Create robust error handling for missing connections or non‑DBConnection types when accessing workbook.DataConnections.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsDemo
{
    // Loads a workbook, accesses its DataConnections collection, verifies the presence of at least one connection, casts the first item to DBConnection, reads the Name property, prints it to the console, and optionally saves the file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the collection of external data connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Ensure there is at least one connection
            if (connections.Count > 0)
            {
                // Retrieve the first connection
                ExternalConnection externalConn = connections[0];

                // Check if the connection is a DBConnection
                if (externalConn is DBConnection dbConn)
                {
                    // Read the Name property of the DBConnection
                    string connectionName = dbConn.Name;

                    // Output the name to the console
                    Console.WriteLine("First DBConnection Name: " + connectionName);
                }
                else
                {
                    Console.WriteLine("The first connection is not a DBConnection.");
                }
            }
            else
            {
                Console.WriteLine("No external connections found in the workbook.");
            }

            // Save the workbook (optional, but follows lifecycle rules)
            workbook.Save("output.xlsx");
        }
    }
}
