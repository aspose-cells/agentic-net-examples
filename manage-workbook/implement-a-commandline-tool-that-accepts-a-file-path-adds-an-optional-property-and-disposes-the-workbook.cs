// Title: C# CLI Tool to Load an Excel Workbook, Optionally Add a Custom Property, and Dispose with Aspose.Cells
// Description: A concise .NET console application that accepts a file path, opens the workbook via Aspose.Cells, adds a custom document property when a name/value pair is supplied, and automatically disposes the Workbook using a `using` block. No file is saved, making it ideal for validation or metadata tagging tasks.
// Keywords: Aspose.Cells C# | CLI workbook loader | add custom document property | Workbook.Dispose | using statement | Excel file command line | batch Excel processing | CI pipeline Excel tag | GitHub Aspose.Cells example | code snippet C# console
// Common Searches: how to open an Excel file with Aspose.Cells in C# console | add custom document property using Aspose.Cells | ensure Workbook.Dispose is called in .NET | C# command line tool for Aspose.Cells | sample code for loading and disposing a workbook
// Developer Intent: Load a workbook from a path, optionally attach a custom property, and guarantee proper disposal.
// Use Cases: Tag a batch of Excel files with metadata before archiving. | Validate that workbooks can be opened and resources released in CI builds. | Create a lightweight utility for automated property injection without persisting changes.
// AI Prompts: Write a C# console program that uses Aspose.Cells to open a workbook from a command‑line path, adds a custom document property if two extra arguments are provided, and disposes the workbook automatically. | Show how to modify the code to save the workbook after adding the property while still using a `using` block for disposal. | Explain the resources released by Aspose.Cells when `Workbook.Dispose` is invoked inside a `using` statement.

using System;
using Aspose.Cells;

namespace AsposeCellsCommandLine
{
    // A concise .NET console application that accepts a file path, opens the workbook via Aspose.Cells, adds a custom document property when a name/value pair is supplied, and automatically disposes the Workbook using a `using` block. No file is saved, making it ideal for validation or metadata tagging tasks.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect at least the workbook file path.
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: AsposeCellsCommandLine <filePath> [propertyName propertyValue]");
                return;
            }

            string filePath = args[0];

            // Load the workbook using the constructor that accepts a file path.
            // This follows the provided create/load rule: Workbook(string file)
            using (Workbook workbook = new Workbook(filePath))
            {
                // If a property name and value are supplied, add a custom document property.
                if (args.Length >= 3)
                {
                    string propertyName = args[1];
                    string propertyValue = args[2];

                    // Add the custom property to the workbook.
                    // The CustomDocumentProperties collection provides an Add method.
                    workbook.CustomDocumentProperties.Add(propertyName, propertyValue);

                    Console.WriteLine($"Added custom property: {propertyName} = {propertyValue}");
                }
                else
                {
                    Console.WriteLine("No custom property specified; workbook loaded only.");
                }

                // No explicit save is required per the task description.
                // The workbook will be disposed automatically by the using statement,
                // which invokes Workbook.Dispose() as defined in the API.
            }

            // At this point the workbook has been disposed.
            Console.WriteLine("Workbook disposed.");
        }
    }
}
