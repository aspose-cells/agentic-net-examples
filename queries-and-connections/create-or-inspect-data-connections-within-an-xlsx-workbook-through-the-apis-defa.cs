using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class DataConnectionDemo
{
    static void Main()
    {
        // Load an existing workbook that may contain external data connections
        string inputPath = "input_with_connections.xlsx";
        Workbook workbook;

        if (File.Exists(inputPath))
        {
            workbook = new Workbook(inputPath);
        }
        else
        {
            workbook = new Workbook();
            Console.WriteLine($"Input file '{inputPath}' not found. Created a new workbook.");
        }

        // Access the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Display the total number of connections
        Console.WriteLine("DataConnections count: " + connections.Count);

        // Iterate through each connection and output key properties
        for (int i = 0; i < connections.Count; i++)
        {
            ExternalConnection conn = connections[i];
            Console.WriteLine($"Connection {i + 1}:");
            Console.WriteLine($"  Name: {conn.Name}");
            Console.WriteLine($"  Id: {conn.Id}");
            Console.WriteLine($"  Source Type: {conn.SourceType}");
            Console.WriteLine($"  Refresh On Load: {conn.RefreshOnLoad}");
        }

        // Example modification: enable RefreshOnLoad for the first connection, if any exist
        if (connections.Count > 0)
        {
            connections[0].RefreshOnLoad = true;
            Console.WriteLine("Set RefreshOnLoad = true for the first connection.");
        }

        // Save the workbook (including any modifications made to connections)
        string outputPath = "output_with_modified_connections.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}