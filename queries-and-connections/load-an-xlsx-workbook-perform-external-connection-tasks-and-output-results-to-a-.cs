using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace ExternalConnectionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook (modify as needed)
            string workbookPath = "input.xlsx";

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(workbookPath);

                // Access the collection of external data connections
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Output the number of connections found
                Console.WriteLine($"DataConnections count: {connections.Count}");

                // Iterate through each connection and display selected properties
                for (int i = 0; i < connections.Count; i++)
                {
                    ExternalConnection conn = connections[i];
                    Console.WriteLine($"--- Connection {i + 1} ---");
                    Console.WriteLine($"Name: {conn.Name}");
                    Console.WriteLine($"Class Type: {conn.ClassType}");
                    Console.WriteLine($"Connection String: {conn.ConnectionString}");
                    Console.WriteLine($"Save Data: {conn.SaveData}");
                    Console.WriteLine($"Only Use Connection File: {conn.OnlyUseConnectionFile}");
                    Console.WriteLine($"Connection File: {conn.ConnectionFile}");
                }

                // Example modification: toggle SaveData for the first connection (if any)
                if (connections.Count > 0)
                {
                    ExternalConnection firstConn = connections[0];
                    firstConn.SaveData = !firstConn.SaveData;
                    Console.WriteLine($"Toggled SaveData for connection '{firstConn.Name}' to {firstConn.SaveData}");
                }

                // Save the workbook to persist any changes made to connections
                workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                // Output any errors to the console (which can be redirected to a TXT file)
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}