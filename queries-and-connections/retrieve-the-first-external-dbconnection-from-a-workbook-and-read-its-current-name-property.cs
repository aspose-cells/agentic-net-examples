using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class RetrieveFirstDbConnectionName
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
            // Get the first connection (index 0) and cast it to DBConnection
            ExternalConnection extConn = connections[0];
            if (extConn is DBConnection dbConn)
            {
                // Read the Name property of the DBConnection
                string connectionName = dbConn.Name;
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

        // Save the workbook (unchanged) to a new file
        workbook.Save("output.xlsx");
    }
}