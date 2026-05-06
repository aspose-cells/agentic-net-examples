using System;
using Aspose.Cells;

namespace AsposeCellsDataConnectionsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize a new workbook (default format is XLSX)
            Workbook workbook = new Workbook();

            // Access the DataConnections collection
            var connections = workbook.DataConnections;

            // Display the number of connections (expected to be 0 for a new workbook)
            Console.WriteLine("DataConnections count: " + connections.Count);

            // If there are any connections, list their names
            for (int i = 0; i < connections.Count; i++)
            {
                Console.WriteLine($"Connection {i + 1}: {connections[i].Name}");
            }

            // Save the workbook using the default XLSX format
            workbook.Save("DataConnectionsProcessed.xlsx");

            Console.WriteLine("Workbook saved as DataConnectionsProcessed.xlsx");
        }
    }
}