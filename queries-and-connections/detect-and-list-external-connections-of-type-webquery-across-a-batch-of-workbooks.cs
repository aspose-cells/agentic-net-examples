// Title: Detect and list WebQuery external data connections in multiple Excel workbooks with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a collection of .xlsx files using Aspose.Cells, iterates each worksheet, and prints the name, type, and URL of every WebQuery external data connection. | Extend the program to write the extracted WebQuery connection details (workbook path, worksheet name, connection name, URL) into a CSV file. | Add fallback logic that safely skips worksheets when the ExternalDataConnections property is unavailable in older Aspose.Cells versions. | Implement robust error handling so the script continues processing remaining workbooks if a file is missing, corrupted, or a connection cannot be read.
// Common Searches: how to retrieve web query URLs from Excel files using Aspose.Cells C# | batch extract external data connections of type WebQuery across many workbooks | Aspose.Cells enumerate ExternalDataConnections collection with reflection | C# script to list WebQuery connections in each worksheet of multiple .xlsx files | skip missing ExternalDataConnections property when using older Aspose.Cells versions
// Tags: Aspose.Cells read WebQuery external connections | C# iterate multiple Excel workbooks Aspose.Cells | reflection access ExternalDataConnections property | export WebQuery details to CSV in C# | handle missing ExternalDataConnections gracefully

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads a list of Excel workbook paths, uses Aspose.Cells to open each file, and via reflection accesses each worksheet's ExternalDataConnections collection. It filters for connections of type 'WebQuery' and outputs the workbook, worksheet, connection name, type, and URL. The code includes checks for missing files, absent properties in older library versions, and per‑connection error handling, making it suitable for batch processing of many workbooks.
class Program
{
    static void Main()
    {
        // List of workbook file paths to process
        var workbookFiles = new List<string>
        {
            @"C:\Workbooks\Book1.xlsx",
            @"C:\Workbooks\Book2.xlsx",
            // Add more file paths as needed
        };

        foreach (var filePath in workbookFiles)
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook (Aspose.Cells handles the file format automatically)
                var workbook = new Workbook(filePath);

                // Iterate through each worksheet to examine its external data connections
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Use reflection to access ExternalDataConnections (may not be available in older versions)
                    var connectionsProp = sheet.GetType().GetProperty("ExternalDataConnections");
                    if (connectionsProp == null)
                    {
                        // Property not available – skip this worksheet
                        continue;
                    }

                    var connectionsObj = connectionsProp.GetValue(sheet) as System.Collections.IEnumerable;
                    if (connectionsObj == null)
                        continue;

                    foreach (var connection in connectionsObj)
                    {
                        try
                        {
                            // Retrieve the connection type via reflection
                            var typeProp = connection.GetType().GetProperty("Type");
                            var typeValue = typeProp?.GetValue(connection)?.ToString();

                            // Filter only WebQuery type connections
                            if (!string.Equals(typeValue, "WebQuery", StringComparison.OrdinalIgnoreCase))
                                continue;

                            var name = connection.GetType().GetProperty("Name")?.GetValue(connection);
                            var connStr = connection.GetType().GetProperty("ConnectionString")?.GetValue(connection);

                            Console.WriteLine($"Workbook: {filePath}");
                            Console.WriteLine($"  Worksheet       : {sheet.Name}");
                            Console.WriteLine($"  Connection Name : {name}");
                            Console.WriteLine($"  Connection Type : {typeValue}");
                            Console.WriteLine($"  Connection URL  : {connStr}");
                            Console.WriteLine();
                        }
                        catch (Exception innerEx)
                        {
                            // Handle any issues retrieving individual connection details
                            Console.WriteLine($"  Error reading a connection in worksheet '{sheet.Name}': {innerEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors (e.g., corrupted file, unsupported format)
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }
    }
}
