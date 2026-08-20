// Title: C# – Handle Invalid External Connection Paths with Aspose.Cells
// Description: Demonstrates loading or creating a workbook, accessing its first DataConnection, assigning an OdcFile path, validating the file, and using try‑catch blocks to capture FileNotFoundException or other errors so the application remains stable.
// Keywords: Aspose.Cells | C# | external connection | OdcFile | FileNotFoundException | error handling | invalid path | DataConnections | Workbook | exception handling
// Common Searches: Aspose.Cells validate OdcFile path | C# catch FileNotFoundException for external connection | how to handle missing ODC file in Aspose.Cells | error handling external connections Aspose.Cells .NET
// Developer Intent: Detect and manage errors when an external connection file path is missing or invalid.
// Use Cases: Notify users when the ODC file referenced by a workbook does not exist. | Prevent saving a workbook that contains broken external connections. | Provide a fallback that creates a new workbook if the source file is absent while still handling connection errors.
// AI Prompts: Generate a C# utility method that checks an Aspose.Cells ExternalConnection OdcFile path and returns a detailed error message if the file is missing. | Refactor the sample to use separate catch blocks for FileNotFoundException and generic Exception, logging each case with timestamps. | Create a reusable class that validates all external connections in a workbook and returns a list of invalid paths.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Demonstrates loading or creating a workbook, accessing its first DataConnection, assigning an OdcFile path, validating the file, and using try‑catch blocks to capture FileNotFoundException or other errors so the application remains stable.
class ExternalConnectionErrorHandlingDemo
{
    static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }

    public static void Run()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Load existing workbook if it exists; otherwise create a new one
        Workbook workbook;
        if (File.Exists(inputPath))
        {
            workbook = new Workbook(inputPath);
        }
        else
        {
            Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
            workbook = new Workbook();
        }

        try
        {
            // Get the collection of external connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Ensure there is at least one connection to work with
            if (connections.Count == 0)
            {
                Console.WriteLine("No external connections found in the workbook.");
                return;
            }

            // Work with the first external connection as an example
            ExternalConnection connection = connections[0];

            // Assign an invalid external connection file path
            string invalidOdcPath = @"Z:\NonExistentFolder\InvalidConnection.odc";
            connection.OdcFile = invalidOdcPath;
            Console.WriteLine($"Assigned OdcFile path: {connection.OdcFile}");

            // Validate the path – if the file does not exist, throw an exception
            if (!File.Exists(connection.OdcFile))
            {
                throw new FileNotFoundException("External connection file not found.", connection.OdcFile);
            }

            // If the path is valid, save the workbook
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            // Capture and display any errors related to the external connection path
            Console.WriteLine($"Error handling external connection: {ex.Message}");
        }
    }
}
