using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExternalConnectionsDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load an existing XLSX workbook that already contains external connections
            // -----------------------------------------------------------------
            // (Assumes "InputWithConnections.xlsx" exists in the executable folder)
            Workbook workbook = new Workbook("InputWithConnections.xlsx");

            // -----------------------------------------------------------------
            // 2. Access the collection of external connections
            // -----------------------------------------------------------------
            ExternalConnectionCollection connections = workbook.DataConnections;

            Console.WriteLine($"Number of external connections: {connections.Count}");

            // -----------------------------------------------------------------
            // 3. Iterate through each connection, display key properties,
            //    and modify a few writable properties for demonstration
            // -----------------------------------------------------------------
            for (int i = 0; i < connections.Count; i++)
            {
                ExternalConnection conn = connections[i];

                Console.WriteLine($"--- Connection {i + 1} ---");
                Console.WriteLine($"Name               : {conn.Name}");
                Console.WriteLine($"Class Type         : {conn.ClassType}");
                Console.WriteLine($"Source Type        : {conn.SourceType}");
                Console.WriteLine($"Connection String  : {conn.ConnectionString}");
                Console.WriteLine($"OdcFile (before)   : {conn.OdcFile}");
                Console.WriteLine($"SaveData (before)  : {conn.SaveData}");
                Console.WriteLine($"OnlyUseConnectionFile (before): {conn.OnlyUseConnectionFile}");

                // Example modifications:
                // Set a dummy ODC file path (write‑able property)
                conn.OdcFile = @"C:\Temp\MyConnection.odc";

                // Ensure that data fetched via this connection is saved with the workbook
                conn.SaveData = true;

                // Force the workbook to always use the external ODC file when refreshing
                conn.OnlyUseConnectionFile = true;

                Console.WriteLine($"OdcFile (after)    : {conn.OdcFile}");
                Console.WriteLine($"SaveData (after)   : {conn.SaveData}");
                Console.WriteLine($"OnlyUseConnectionFile (after): {conn.OnlyUseConnectionFile}");
            }

            // -----------------------------------------------------------------
            // 4. Optionally, add a new external link to another workbook (not a connection)
            // -----------------------------------------------------------------
            // Demonstrates using the ExternalLinkCollection.Add method
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            string externalFile = "ExternalData.xlsx";
            string[] sheetNames = new string[] { "Sheet1!A1:B10" };
            int linkIndex = externalLinks.Add(externalFile, sheetNames);
            Console.WriteLine($"Added external link at index {linkIndex}, total links: {externalLinks.Count}");

            // -----------------------------------------------------------------
            // 5. Save the workbook with the modified connection settings
            // -----------------------------------------------------------------
            workbook.Save("OutputWithModifiedConnections.xlsx");

            Console.WriteLine("Workbook saved as 'OutputWithModifiedConnections.xlsx'.");
        }
    }
}