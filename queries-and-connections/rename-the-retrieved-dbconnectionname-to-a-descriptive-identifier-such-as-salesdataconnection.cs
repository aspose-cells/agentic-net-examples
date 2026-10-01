// Title: Rename the first workbook data connection to a custom name (e.g., SalesDataConnection) using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells C# API to change the Name property of the first DataConnection in an Excel workbook to "SalesDataConnection" and save the file. | Programmatically update a workbook's DB connection identifier with Aspose.Cells, then persist the changes to a new .xlsx file.
// Common Searches: C# Aspose.Cells how to change the name of a data connection in an existing Excel file | rename Excel workbook data connection programmatically with Aspose.Cells .NET | set custom identifier for first DB connection in workbook using Aspose.Cells API | Aspose.Cells update DataConnection.Name and save workbook example
// Tags: aspocells rename dataconnection | c# aspocells set connection name | excel workbook data connection update | aspocells modify dbconnection identifier | save workbook after connection rename aspocells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, checks for existing data connections, renames the first connection to "SalesDataConnection", saves the modified workbook to a new file, and includes basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: No worksheets found in the workbook.");
                return;
            }

            // Rename the first data connection if it exists (Workbook-level collection)
            if (workbook.DataConnections.Count > 0)
            {
                workbook.DataConnections[0].Name = "SalesDataConnection";
                Console.WriteLine("Database connection renamed successfully.");
            }
            else
            {
                Console.WriteLine("No data connections found in the workbook.");
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
