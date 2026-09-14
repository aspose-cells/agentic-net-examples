// Title: C# command‑line program to load an Excel workbook with Aspose.Cells, add a custom document property, and ensure proper disposal
// AI Prompts: Generate a C# console application that accepts an Excel file path argument, opens the file with Aspose.Cells.Workbook, adds a custom document property named "OptionalProperty" with the value "OptionalValue", and disposes the workbook in a finally block. | Extend the console app to allow the property name and value to be supplied as optional command‑line parameters, defaulting to "OptionalProperty" and "OptionalValue" when omitted, while still guaranteeing workbook disposal. | Implement robust error handling for the console tool that logs any exceptions during workbook loading or property addition and ensures the Workbook object is always disposed, even on failure.
// Common Searches: how to add a custom document property to an Excel file using Aspose.Cells in a C# console application | C# command line tool for loading an Excel workbook with Aspose.Cells and releasing resources | example of disposing Aspose.Cells Workbook safely in .NET console app | Aspose.Cells command line program to set optional property and close workbook | using Aspose.Cells to add custom properties from command line arguments in C#
// Tags: aspocells add custom document property | aspocells workbook disposal .net | c# console load excel aspocells | command line excel processing aspocells | optional custom property workbook aspocells

using System;
using Aspose.Cells;

namespace AsposeCellsCommandLine
{
    // A C# console utility that reads an Excel file path from the command line, loads the workbook with Aspose.Cells, adds (or overwrites) a custom document property called "OptionalProperty" with a default value "OptionalValue", handles any loading errors, and guarantees the Workbook is disposed to free resources.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a file path argument was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsCommandLine <excel-file-path>");
                return;
            }

            string filePath = args[0];

            // Load the workbook from the specified file
            Workbook workbook = null;
            try
            {
                workbook = new Workbook(filePath);

                // Add an optional custom document property
                // If the property already exists, it will be overwritten
                workbook.CustomDocumentProperties.Add("OptionalProperty", "OptionalValue");

                // Optionally, you could perform additional operations here
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
            finally
            {
                // Ensure the workbook is properly disposed to release resources
                workbook?.Dispose();
            }
        }
    }
}
