using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Path to the existing XLSX workbook
        string inputPath = "input.xlsx";

        // Load the workbook from the file (uses the provided constructor)
        Workbook workbook = new Workbook(inputPath);

        // Retrieve the collection of external data connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        Console.WriteLine($"DataConnections count: {connections.Count}");

        // Iterate through each connection and apply modifications
        for (int i = 0; i < connections.Count; i++)
        {
            ExternalConnection conn = connections[i];
            Console.WriteLine($"Connection {i + 1}: {conn.Name}");

            // Example modification: always use the connection file when refreshing
            conn.OnlyUseConnectionFile = true;

            // If the connection is a WebQueryConnection, set additional property
            if (conn is WebQueryConnection webConn)
            {
                // Import XML source data instead of HTML table
                webConn.IsXmlSourceData = true;
                Console.WriteLine($"WebQueryConnection IsXmlSourceData set to {webConn.IsXmlSourceData}");
            }
        }

        // Ensure that all connections are refreshed when the workbook is opened in Excel
        workbook.Worksheets.IsRefreshAllConnections = true;

        // Save the modified workbook (uses the provided Save method)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}