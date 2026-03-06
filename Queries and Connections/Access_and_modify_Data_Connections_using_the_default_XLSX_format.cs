using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class DataConnectionsDemo
{
    static void Main()
    {
        // Create a new workbook (default XLSX format)
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "DataSheet";
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");

        // Set the workbook to refresh all connections when opened in Excel
        workbook.Worksheets.IsRefreshAllConnections = true;

        // Access the collection of external data connections
        ExternalConnectionCollection connections = workbook.DataConnections;
        Console.WriteLine("Initial DataConnections count: " + connections.Count);

        // If any connections exist, modify several of their properties
        for (int i = 0; i < connections.Count; i++)
        {
            ExternalConnection conn = connections[i];
            Console.WriteLine($"Connection {i + 1} Name: {conn.Name}");

            // Toggle the OnlyUseConnectionFile flag
            conn.OnlyUseConnectionFile = !conn.OnlyUseConnectionFile;
            Console.WriteLine($"OnlyUseConnectionFile set to: {conn.OnlyUseConnectionFile}");

            // Assign a sample ODC file path
            conn.OdcFile = @"C:\Temp\MyConnection.odc";
            Console.WriteLine($"OdcFile set to: {conn.OdcFile}");

            // Ensure that data fetched via the connection is saved with the workbook
            conn.SaveData = true;
            Console.WriteLine($"SaveData set to: {conn.SaveData}");
        }

        // Save the workbook to disk in XLSX format
        workbook.Save("DataConnectionsDemo.xlsx");
    }
}