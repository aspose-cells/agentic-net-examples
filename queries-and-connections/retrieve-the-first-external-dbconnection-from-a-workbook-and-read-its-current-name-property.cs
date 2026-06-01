using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        if (connections.Count > 0)
        {
            // Get the first connection in the collection
            ExternalConnection connection = connections[0];

            // If the connection is a DBConnection, read its Name property
            if (connection is DBConnection dbConnection)
            {
                Console.WriteLine("First DBConnection Name: " + dbConnection.Name);
            }
            else
            {
                // Fallback: display the name of the first connection regardless of its type
                Console.WriteLine("First connection is not a DBConnection. Name: " + connection.Name);
            }
        }
        else
        {
            Console.WriteLine("No external connections found in the workbook.");
        }

        // Save the workbook (required lifecycle step, even if unchanged)
        workbook.Save("output.xlsx");
    }
}