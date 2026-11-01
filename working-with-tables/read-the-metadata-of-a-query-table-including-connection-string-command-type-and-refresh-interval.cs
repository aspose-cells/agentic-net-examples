// Title: C# example to read connection string, command text, command type, and refresh interval of an Excel query table using Aspose.Cells
// Common Searches: how to read the connection string of a query table in an Excel workbook with Aspose.Cells C# | sample code for obtaining command text and command type from an Excel query table using Aspose.Cells .NET | retrieve refresh interval of a query table programmatically via Aspose.Cells | using reflection to safely access query table properties in Aspose.Cells for .NET | Aspose.Cells example to list query table metadata from a worksheet
// Tags: Aspose.Cells query table metadata extraction | C# retrieve query table connection string | Aspose.Cells get query table command type | Excel query table refresh interval access | reflection based property reading Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample loads an Excel workbook with Aspose.Cells, checks for a query table on the first worksheet, and uses reflection to obtain the Connection/ConnectionString, Command/CommandText, CommandType, and RefreshInterval properties of the first QueryTable, printing the values to the console.
class QueryTableMetadataReader
{
    static void Main()
    {
        try
        {
            // Path to the workbook
            string filePath = "Sample.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one query table
            if (sheet.QueryTables.Count == 0)
            {
                Console.WriteLine("No query tables found in the worksheet.");
                return;
            }

            // Get the first query table
            QueryTable queryTable = sheet.QueryTables[0];

            // Retrieve metadata using reflection to handle possible API variations
            string connection = GetPropertyValue(queryTable, "Connection") ??
                                GetPropertyValue(queryTable, "ConnectionString") ??
                                "N/A";

            string command = GetPropertyValue(queryTable, "Command") ??
                             GetPropertyValue(queryTable, "CommandText") ??
                             "N/A";

            string commandType = GetPropertyValue(queryTable, "CommandType") ?? "N/A";

            string refreshInterval = GetPropertyValue(queryTable, "RefreshInterval") ?? "N/A";

            // Output the metadata
            Console.WriteLine("Query Table Metadata:");
            Console.WriteLine($"Connection String : {connection}");
            Console.WriteLine($"Command           : {command}");
            Console.WriteLine($"Command Type      : {commandType}");
            Console.WriteLine($"Refresh Interval  : {refreshInterval} minutes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to safely obtain property values via reflection
    private static string GetPropertyValue(object obj, string propertyName)
    {
        PropertyInfo propInfo = obj.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
        if (propInfo != null)
        {
            object value = propInfo.GetValue(obj);
            return value?.ToString();
        }
        return null;
    }
}
