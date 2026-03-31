using System;
using System.IO;
using Aspose.Cells;

class ExternalConnectionsDemo
{
    static void Main()
    {
        string inputPath = "InputWithConnections.xlsx";
        Workbook workbook;

        if (File.Exists(inputPath))
        {
            workbook = new Workbook(inputPath);
        }
        else
        {
            workbook = new Workbook(); // create a new workbook if the file is missing
        }

        // Access the collection of external connections
        var connections = workbook.DataConnections;
        Console.WriteLine("External connections count: " + connections.Count);

        if (connections.Count > 0)
        {
            // Work with the first connection in the collection
            var conn = connections[0];

            // Display original property values
            Console.WriteLine("Original Name: " + conn.Name);
            Console.WriteLine("Original OdcFile: " + conn.OdcFile);
            Console.WriteLine("Original SaveData: " + conn.SaveData);
            Console.WriteLine("Original OnlyUseConnectionFile: " + conn.OnlyUseConnectionFile);

            // Modify properties to demonstrate writing
            conn.OdcFile = @"C:\Connections\MyConnection.odc";
            conn.SaveData = true;
            conn.OnlyUseConnectionFile = true;

            // Display modified property values
            Console.WriteLine("Modified OdcFile: " + conn.OdcFile);
            Console.WriteLine("Modified SaveData: " + conn.SaveData);
            Console.WriteLine("Modified OnlyUseConnectionFile: " + conn.OnlyUseConnectionFile);
        }
        else
        {
            Console.WriteLine("No external connections found in the workbook.");
        }

        // Demonstrate adding an external link to another workbook
        var externalLinks = workbook.Worksheets.ExternalLinks;
        int linkIndex = externalLinks.Add("LinkedWorkbook.xlsx", new string[] { "Sheet1", "Sheet2" });
        Console.WriteLine("Added external link at index: " + linkIndex);
        Console.WriteLine("External link data source: " + externalLinks[linkIndex].DataSource);

        // Save the workbook with the updated connection settings and external link
        workbook.Save("OutputWithModifiedConnections.xlsx");
    }
}