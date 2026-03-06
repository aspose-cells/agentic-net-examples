using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class ODataConnectionInfoDemo
    {
        public static void Run()
        {
            // Load an existing XLSX workbook that contains OData connections
            // (Replace the file path with the actual location of your workbook)
            Workbook workbook = new Workbook("ODataSample.xlsx");

            // Access the collection of external data connections in the workbook
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Iterate through each connection and display OData related information
            foreach (ExternalConnection conn in connections)
            {
                // Show basic connection details
                Console.WriteLine($"Connection Name: {conn.Name}");
                Console.WriteLine($"Source Type: {conn.SourceType}");

                // For OData connections the SourceType is typically DataFeedDataModel (value 101)
                // or WebQuery (value 4) depending on how the connection was created.
                // Display the connection string which holds the OData endpoint URL and parameters.
                Console.WriteLine($"Connection String: {conn.ConnectionString}");

                // Additional useful properties
                Console.WriteLine($"Refresh On Load: {conn.RefreshOnLoad}");
                Console.WriteLine($"Background Refresh: {conn.BackgroundRefresh}");
                Console.WriteLine(new string('-', 40));
            }

            // Optionally, modify a property (e.g., enable refresh on load) for demonstration
            if (connections.Count > 0)
            {
                connections[0].RefreshOnLoad = true;
            }

            // Save the workbook back to the default XLSX format
            workbook.Save("ODataSample_Updated.xlsx");
        }
    }
}