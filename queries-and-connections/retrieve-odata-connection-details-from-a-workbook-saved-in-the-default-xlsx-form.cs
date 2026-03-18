using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsODataRetrieval
{
    public class Program
    {
        public static void Main()
        {
            // Load the existing XLSX workbook (default format)
            string workbookPath = "input.xlsx";
            Workbook workbook = new Workbook(workbookPath);

            // Access the collection of external data connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Iterate through each connection and output OData related details
            for (int i = 0; i < connections.Count; i++)
            {
                ExternalConnection conn = connections[i];

                // Common properties useful for OData connections
                Console.WriteLine($"Connection #{i + 1}");
                Console.WriteLine($"Name               : {conn.Name}");
                Console.WriteLine($"ClassType          : {conn.ClassType}");
                Console.WriteLine($"ConnectionString   : {conn.ConnectionString}");
                Console.WriteLine($"OdcFile            : {conn.OdcFile}");
                Console.WriteLine($"RefreshOnLoad      : {conn.RefreshOnLoad}");
                Console.WriteLine($"SavePassword       : {conn.SavePassword}");
                Console.WriteLine(new string('-', 40));
            }

            // No need to save the workbook as we only read connection details
        }
    }
}