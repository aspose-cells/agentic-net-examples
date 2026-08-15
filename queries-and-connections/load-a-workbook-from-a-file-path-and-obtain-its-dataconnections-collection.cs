// Title: C# – Load an Excel workbook from a file path and access its DataConnections with Aspose.Cells
// Description: Demonstrates how to verify an Excel file’s existence, load it into an Aspose.Cells Workbook, retrieve the Workbook.DataConnections collection, and output the total number of external data connections while handling possible exceptions.
// Keywords: Aspose.Cells | C# | load workbook from file | Workbook.DataConnections | external data connections | Excel file existence check | count data connections | error handling Aspose.Cells
// Common Searches: Aspose.Cells load workbook and list data connections | How to get DataConnections count in C# Excel file | Retrieve external data connections with Aspose.Cells .NET | Check if Excel file exists before loading Aspose.Cells | C# code to access Workbook.DataConnections
// Developer Intent: Load a workbook from a specified path and obtain its DataConnections collection.
// Use Cases: Validate that an incoming Excel file contains the required number of external data connections before processing. | Log the count of data connections for auditing or troubleshooting data import pipelines. | Iterate over workbook.DataConnections to inspect each connection’s name, type, or connection string.
// AI Prompts: Generate C# code that checks for an Excel file, loads it with Aspose.Cells, and prints the DataConnections count. | Show how to enumerate workbook.DataConnections and display each connection’s properties in C#. | Explain best practices for exception handling when accessing DataConnections after loading a workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to verify an Excel file’s existence, load it into an Aspose.Cells Workbook, retrieve the Workbook.DataConnections collection, and output the total number of external data connections while handling possible exceptions.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Get the collection of external data connections
            var dataConnections = workbook.DataConnections;

            // Display the number of data connections present
            Console.WriteLine($"DataConnections count: {dataConnections.Count}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
