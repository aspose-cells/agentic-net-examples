using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main(string[] args)
    {
        // Path to the source workbook (replace with your actual file)
        string workbookPath = "input.xlsx";

        // Path to the output text file that will contain the connection names
        string outputPath = "ExternalConnectionNames.txt";

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Get the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Write each connection name to the text file
        using (StreamWriter writer = new StreamWriter(outputPath))
        {
            for (int i = 0; i < connections.Count; i++)
            {
                // Each connection has a unique Name property
                writer.WriteLine(connections[i].Name);
            }
        }

        Console.WriteLine($"Exported {connections.Count} external connection name(s) to '{outputPath}'.");
    }
}