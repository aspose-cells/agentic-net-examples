// Title: Load an Excel workbook from a file path and list its external data connections using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a specified .xlsx file with Aspose.Cells and prints each external connection name. | Show how to retrieve the DataConnections collection from a Workbook object and display the connection count. | Provide a snippet that iterates over workbook.DataConnections and writes connection details to the console.
// Common Searches: Aspose.Cells C# get list of external data connections from an existing Excel file | How to read DataConnections collection of a workbook using Aspose.Cells for .NET | C# code sample for enumerating external connections in an .xlsx with Aspose.Cells | Retrieve and display connection names from Excel workbook using Aspose.Cells API
// Tags: load workbook from file Aspose.Cells C# | enumerate DataConnections collection Aspose.Cells | list external Excel connections .NET | Aspose.Cells workbook.DataConnections usage | print connection names Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Demonstrates loading an Excel file via Aspose.Cells, accessing the workbook's DataConnections collection, outputting the total count, and iterating to display each connection's name, with a note on saving changes if needed.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the specified file (uses Workbook(string) constructor)
        Workbook workbook = new Workbook(filePath);

        // Get the collection of external data connections in the workbook
        ExternalConnectionCollection dataConnections = workbook.DataConnections;

        // Display the number of data connections found
        Console.WriteLine($"DataConnections count: {dataConnections.Count}");

        // Iterate through the connections (if any) and output their names
        for (int i = 0; i < dataConnections.Count; i++)
        {
            ExternalConnection connection = dataConnections[i];
            Console.WriteLine($"Connection {i + 1}: {connection.Name}");
        }

        // If you modify connections and need to persist changes, you could save:
        // workbook.Save("output.xlsx");
    }
}
