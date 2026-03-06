using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace ExternalConnectionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel workbook
            string excelPath = "input.xlsx";

            // Load the workbook using the provided constructor (Workbook(string))
            Workbook workbook = new Workbook(excelPath);

            // Access the collection of external connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Prepare a StringWriter to collect output for the text file
            using (StringWriter txtWriter = new StringWriter())
            {
                // Write header information to console and text output
                Console.WriteLine($"External connections count: {connections.Count}");
                txtWriter.WriteLine($"External connections count: {connections.Count}");

                // Iterate through each external connection and display its properties
                for (int i = 0; i < connections.Count; i++)
                {
                    ExternalConnection conn = connections[i];

                    // Gather details
                    string name = conn.Name ?? "N/A";
                    string classType = conn.ClassType.ToString();
                    string connectionString = conn.ConnectionString ?? "N/A";

                    // Output to console
                    Console.WriteLine($"Connection {i + 1}:");
                    Console.WriteLine($"  Name            : {name}");
                    Console.WriteLine($"  Class Type      : {classType}");
                    Console.WriteLine($"  Connection String: {connectionString}");

                    // Append the same information to the text writer
                    txtWriter.WriteLine($"Connection {i + 1}:");
                    txtWriter.WriteLine($"  Name            : {name}");
                    txtWriter.WriteLine($"  Class Type      : {classType}");
                    txtWriter.WriteLine($"  Connection String: {connectionString}");
                }

                // Write the collected information to a TXT file
                string txtPath = "ExternalConnectionsInfo.txt";
                File.WriteAllText(txtPath, txtWriter.ToString());

                Console.WriteLine($"Connection details have been written to '{txtPath}'.");
            }

            // Dispose the workbook (optional, as it implements IDisposable)
            workbook.Dispose();
        }
    }
}