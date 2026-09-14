// Title: Read the Title built‑in document property from an Excel workbook using Aspose.Cells in C#
// AI Prompts: Generate C# code that opens a given .xlsx file with Aspose.Cells, verifies the file exists, and prints the workbook's built‑in Title property. | Show how to access the BuiltInDocumentProperties of a Workbook and extract the Title value while handling possible exceptions. | Provide a snippet that loads an Excel file with Aspose.Cells and outputs its Title metadata to the console.
// Common Searches: c# aspocells get title built-in document property from excel file | how to read built-in document properties like Title using Aspose.Cells .NET | example code to retrieve Excel workbook Title with Aspose.Cells and handle missing file | Aspose.Cells read Title property from .xlsx workbook in C# | retrieve Excel file metadata Title using Aspose.Cells and exception handling
// Tags: Aspose.Cells read built‑in document properties | C# retrieve Excel workbook Title property | load .xlsx workbook and access BuiltInDocumentProperties | exception handling for Aspose.Cells workbook loading | verify Excel file existence with Aspose.Cells | output workbook Title metadata to console

using System;
using System.IO;
using Aspose.Cells;

// The example loads 'input.xlsx' with Aspose.Cells, checks for the file's existence, accesses the BuiltInDocumentProperties collection, reads the Title property, prints it to the console, and gracefully handles any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file '{inputFile}' was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputFile);

            // Access the built‑in document properties collection
            var properties = workbook.BuiltInDocumentProperties;

            // Retrieve the Title property
            string title = properties.Title;

            // Output the title to verify
            Console.WriteLine("Title: " + title);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
