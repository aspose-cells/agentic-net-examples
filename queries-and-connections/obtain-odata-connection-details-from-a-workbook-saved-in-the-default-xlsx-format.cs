using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class ODataConnectionInfo
{
    static void Main()
    {
        // Path to the existing XLSX workbook
        string workbookPath = "sample.xlsx";

        // Load the workbook (uses the Workbook(string) constructor)
        Workbook workbook = new Workbook(workbookPath);

        // Iterate through all external data connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Basic connection details (available for all connection types)
            Console.WriteLine($"Connection Name   : {connection.Name}");
            Console.WriteLine($"Connection Type   : {connection.GetType().Name}");
            Console.WriteLine($"Connection String : {connection.ConnectionString}");
            Console.WriteLine($"Connection File   : {connection.ConnectionFile}");

            // Identify OData connections by type name (Aspose.Cells may have a specific ODataConnection class)
            if (connection.GetType().Name.Contains("OData", StringComparison.OrdinalIgnoreCase))
            {
                // OData‑specific properties can be accessed here if needed.
                // For demonstration, we simply note that this is an OData connection.
                Console.WriteLine("-> This is an OData connection.");
            }

            Console.WriteLine(new string('-', 40));
        }

        // Clean up
        workbook.Dispose();
    }
}