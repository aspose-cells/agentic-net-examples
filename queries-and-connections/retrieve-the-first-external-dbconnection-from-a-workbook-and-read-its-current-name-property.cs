// Title: Retrieve the name of the first external database connection from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get name of first external data connection in Excel file | How to access external DB connections collection in a workbook with Aspose.Cells .NET | Read external database connection properties from a worksheet using Aspose.Cells C# | Retrieve external data connection name via reflection Aspose.Cells | List external connections in an Excel workbook with Aspose.Cells API
// Tags: Aspose.Cells external data connection name | C# read external DB connection from workbook | Aspose.Cells get first external connection | Excel workbook external connections C# | reflection access ExternalDataConnections Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through its worksheets, and uses reflection to obtain the worksheet's ExternalDataConnections collection. It checks the collection count, retrieves the first connection, reads its Name property, and prints the connection name, handling missing files or reflection errors gracefully.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);
            bool connectionFound = false;

            // Iterate through worksheets to locate external DB connections using reflection
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Attempt to get the ExternalDataConnections property via reflection
                    var connectionsProp = sheet.GetType().GetProperty("ExternalDataConnections");
                    if (connectionsProp == null) continue;

                    var connections = connectionsProp.GetValue(sheet);
                    if (connections == null) continue;

                    // Get the Count property of the collection
                    var countProp = connections.GetType().GetProperty("Count");
                    if (countProp == null) continue;

                    int count = (int)countProp.GetValue(connections);
                    if (count <= 0) continue;

                    // Retrieve the first connection (index 0)
                    var indexer = connections.GetType().GetMethod("get_Item", new[] { typeof(int) });
                    if (indexer == null) continue;

                    var firstConnection = indexer.Invoke(connections, new object[] { 0 });
                    if (firstConnection == null) continue;

                    // Get the Name property of the connection
                    var nameProp = firstConnection.GetType().GetProperty("Name");
                    string connectionName = nameProp?.GetValue(firstConnection)?.ToString() ?? "Unnamed";

                    Console.WriteLine($"First external DB connection name in sheet '{sheet.Name}': {connectionName}");
                    connectionFound = true;
                    break; // Stop after finding the first connection
                }
                catch (Exception innerEx)
                {
                    // Handle reflection-related errors for the current worksheet
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {innerEx.Message}");
                }
            }

            if (!connectionFound)
            {
                Console.WriteLine("No external DB connections found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
