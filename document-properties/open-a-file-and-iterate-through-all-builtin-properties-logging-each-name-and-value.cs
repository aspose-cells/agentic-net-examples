// Title: Read and list all built‑in document properties of an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that checks for an .xlsx file, loads it with Aspose.Cells, and prints each built‑in document property (name and value) to the console. | Create a method that opens a workbook with Aspose.Cells, iterates over its BuiltInDocumentProperties collection, and logs the property details while handling missing‑file and load exceptions.
// Common Searches: C# Aspose.Cells list built‑in document properties of an Excel workbook | sample code to read Excel file metadata using Aspose.Cells .NET | how to handle missing Excel file when loading with Aspose.Cells in C# | Aspose.Cells example for iterating workbook metadata properties | retrieve author, title, and other built‑in properties from .xlsx with Aspose.Cells
// Tags: Aspose.Cells enumerate workbook metadata | C# read Excel workbook properties with Aspose.Cells | log property name and value Aspose.Cells | validate Excel file existence before loading | exception handling for Aspose.Cells workbook load

using System;
using System.IO;
using Aspose.Cells;

// The program checks that the specified Excel file exists, loads it into an Aspose.Cells Workbook, iterates through the BuiltInDocumentProperties collection, and writes each property's name and value to the console, with error handling for missing files and load failures.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook
            var workbook = new Workbook(inputPath);

            // Iterate through all built‑in document properties
            foreach (var prop in workbook.BuiltInDocumentProperties)
            {
                // Output property name and its value
                Console.WriteLine($"{prop.Name}: {prop.Value}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
