using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class ExportExternalConnectionNames
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Path of the plain text file that will contain the connection names
        string outputFile = "ExternalConnectionNames.txt";

        // Open a StreamWriter to write each connection name on a separate line
        using (StreamWriter writer = new StreamWriter(outputFile))
        {
            // Get the collection of external connections from the workbook
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Iterate through the collection and write each name
            for (int i = 0; i < connections.Count; i++)
            {
                writer.WriteLine(connections[i].Name);
            }
        }

        // Optional: inform the user that the operation completed
        Console.WriteLine($"Exported {workbook.DataConnections.Count} connection name(s) to '{outputFile}'.");
    }
}