using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Array of workbook file paths to be examined
        string[] workbookFiles = new string[]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx"
            // Add additional file paths as needed
        };

        foreach (string filePath in workbookFiles)
        {
            // Load the workbook (lifecycle rule)
            Workbook workbook = new Workbook(filePath);

            Console.WriteLine($"Processing workbook: {filePath}");

            // Access the external connections collection (rule)
            ExternalConnectionCollection connections = workbook.DataConnections;

            bool anyWebQuery = false;

            // Iterate through all external connections
            for (int i = 0; i < connections.Count; i++)
            {
                // Retrieve connection by index (rule)
                ExternalConnection conn = connections[i];

                // Determine if this connection is a WebQueryConnection
                if (conn is WebQueryConnection webConn)
                {
                    anyWebQuery = true;
                    Console.WriteLine($"  WebQuery Connection #{i + 1}");
                    Console.WriteLine($"    Name           : {webConn.Name}");
                    Console.WriteLine($"    URL            : {webConn.Url}");
                    Console.WriteLine($"    IsHtmlTables   : {webConn.IsHtmlTables}");
                    Console.WriteLine($"    ClassType      : {webConn.ClassType}");
                }
            }

            if (!anyWebQuery)
            {
                Console.WriteLine("  No WebQuery connections found.");
            }

            // Save the workbook unchanged (lifecycle rule)
            string outputPath = $"Processed_{System.IO.Path.GetFileName(filePath)}";
            workbook.Save(outputPath);
        }
    }
}