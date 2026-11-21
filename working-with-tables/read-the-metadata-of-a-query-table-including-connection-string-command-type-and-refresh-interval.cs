// Title: How to read query table properties (name, source data, destination cell, refresh period, refresh on load) from an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get query table source data and destination cell from workbook | Read refresh period of Excel query tables with Aspose.Cells .NET API | Enumerate all query tables in a worksheet and display their metadata using Aspose.Cells | C# retrieve query table name and refresh on load flag from Excel file with Aspose.Cells | How to safely access query table properties that may not exist in older Aspose.Cells versions
// Tags: Aspose.Cells query table metadata extraction | C# read Excel query table properties | retrieve query table refresh interval Aspose.Cells | list query tables in worksheet Aspose.Cells | dynamic property handling Aspose.Cells objects

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, accesses the first worksheet, obtains its QueryTableCollection, and iterates through each QueryTable. Using a generic dynamic helper it safely reads properties such as Name, DestinationCell, SourceData, RefreshPeriod, and RefreshOnLoad, then prints the values to the console while handling missing properties and exceptions.
class QueryTableMetadataReader
{
    static void Main()
    {
        // Path to the input workbook
        string workbookPath = "InputWorkbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of query tables in the worksheet
            QueryTableCollection queryTables = sheet.QueryTables;

            // Iterate through each query table and read its metadata
            foreach (QueryTable qt in queryTables)
            {
                // Use dynamic to safely access properties that may differ between versions
                dynamic dynQt = qt;

                string name = GetProperty(dynQt, "Name", "N/A");
                string destination = GetProperty(dynQt, "DestinationCell", "N/A");
                string sourceData = GetProperty(dynQt, "SourceData", "N/A");
                int refreshPeriod = GetProperty(dynQt, "RefreshPeriod", 0);
                bool refreshOnLoad = GetProperty(dynQt, "RefreshOnLoad", false);

                // Output the metadata
                Console.WriteLine($"Query Table Name: {name}");
                Console.WriteLine($"Destination Cell: {destination}");
                Console.WriteLine($"Source Data: {sourceData}");
                Console.WriteLine($"Refresh Period (minutes): {refreshPeriod}");
                Console.WriteLine($"Refresh On Load: {refreshOnLoad}");
                Console.WriteLine(new string('-', 40));
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine("An error occurred while reading the workbook:");
            Console.WriteLine(ex.Message);
        }
    }

    // Generic helper to retrieve a property value safely; returns defaultValue if unavailable
    private static T GetProperty<T>(dynamic obj, string propertyName, T defaultValue)
    {
        try
        {
            var value = ((object)obj).GetType().GetProperty(propertyName)?.GetValue(obj, null);
            if (value is T typedValue)
                return typedValue;

            // Attempt conversion for common types (e.g., CellArea to string)
            if (typeof(T) == typeof(string) && value != null)
                return (T)(object)value.ToString();

            return defaultValue;
        }
        catch
        {
            return defaultValue;
        }
    }
}
