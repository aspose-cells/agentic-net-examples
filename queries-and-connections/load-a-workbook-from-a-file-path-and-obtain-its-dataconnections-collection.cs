using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "example.xlsx";

        // Load the workbook from the specified file (uses Workbook(string) constructor)
        Workbook workbook = new Workbook(filePath);

        // Obtain the collection of external data connections
        ExternalConnectionCollection dataConnections = workbook.DataConnections;

        // Display the number of connections and their names
        Console.WriteLine("DataConnections count: " + dataConnections.Count);
        for (int i = 0; i < dataConnections.Count; i++)
        {
            ExternalConnection connection = dataConnections[i];
            Console.WriteLine($"Connection {i + 1}: {connection.Name}");
        }
    }
}