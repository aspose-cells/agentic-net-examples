// Title: Copy all custom document properties from one Excel workbook to another using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens a source .xlsx file, creates a new workbook, copies every custom document property from the source to the new workbook, and saves the result to a given path. | Generate a method that iterates over Workbook.CustomDocumentProperties, checks if each property exists in the destination workbook, updates existing ones or adds missing ones, and includes error handling for a missing source file.
// Common Searches: how to copy custom document properties from one Excel file to another using Aspose.Cells C# | Aspose.Cells C# copy all custom workbook metadata to a new workbook | programmatically transfer Excel custom properties with Aspose.Cells .NET | C# example for cloning custom document properties between workbooks using Aspose.Cells
// Tags: Aspose.Cells replicate custom document properties | C# clone Excel workbook metadata | transfer custom properties between workbooks Aspose.Cells | programmatic Excel .xlsx property duplication .NET | add or update custom properties Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads source.xlsx, creates an empty destination workbook, iterates through sourceWorkbook.CustomDocumentProperties, adds or updates each property in destinationWorkbook.CustomDocumentProperties, and saves the destination as destination.xlsx while handling missing files and exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string destinationPath = "destination.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create a new (empty) destination workbook
                Workbook destinationWorkbook = new Workbook();

                // Copy all custom document properties from source to destination
                foreach (var prop in sourceWorkbook.CustomDocumentProperties)
                {
                    // If the property already exists, update its value
                    if (destinationWorkbook.CustomDocumentProperties.Contains(prop.Name))
                    {
                        destinationWorkbook.CustomDocumentProperties[prop.Name].Value = prop.Value;
                    }
                    else
                    {
                        // Add property; convert value to string as Add expects a string argument
                        string valueAsString = prop.Value?.ToString() ?? string.Empty;
                        destinationWorkbook.CustomDocumentProperties.Add(prop.Name, valueAsString);
                    }
                }

                // Save the destination workbook
                destinationWorkbook.Save(destinationPath);
                Console.WriteLine($"Destination workbook saved to {destinationPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
