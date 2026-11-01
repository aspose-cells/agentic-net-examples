// Title: Load an Excel workbook from a file path and list its data connections using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a workbook from a specified path with Aspose.Cells, verifies the file exists, and prints each data connection's name and type. | Show how to obtain the DataConnections collection from a loaded workbook and safely iterate over it in a .NET application.
// Common Searches: aspnet cells get list of data connections from existing .xlsx file c# | c# aspose.cells enumerate workbook data connections after loading file | how to check workbook file exists before accessing DataConnections in Aspose.Cells | retrieve connection name and type from Excel workbook using Aspose.Cells .NET
// Tags: load workbook and access DataConnections Aspose.Cells | enumerate Excel data connections C# | ensure workbook file exists prior to loading Aspose.Cells | display data connection details Aspose.Cells | handle DataConnections collection errors Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// // Loads an Excel workbook from a given path, confirms the file is present, accesses the workbook's DataConnections collection, and outputs each connection's name and type while handling potential runtime exceptions.
class Program
{
    static void Main()
    {
        // Path to the workbook file
        string filePath = @"C:\Path\To\Your\Workbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Get the collection of data connections in the workbook
            var dataConnections = workbook.DataConnections;

            // Iterate through the connections and display their names and types
            foreach (var connection in dataConnections)
            {
                Console.WriteLine($"Connection Name: {connection.Name}, Type: {connection.Type}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
