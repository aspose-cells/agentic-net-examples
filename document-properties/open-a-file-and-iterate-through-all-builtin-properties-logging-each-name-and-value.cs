// Title: Read and log all built‑in document properties from an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# code to enumerate built‑in Excel document properties using Aspose.Cells | how to list all workbook metadata properties with Aspose.Cells .NET | example of reading Excel built‑in properties and printing them in C# | Aspose.Cells iterate BuiltInDocumentProperties collection tutorial | retrieve and display Excel file properties programmatically with Aspose.Cells
// Tags: Aspose.Cells retrieve built‑in properties | C# read Excel document metadata | log property name and value Aspose | traverse BuiltInDocumentProperties collection | extract Excel built‑in metadata programmatically

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for an input.xlsx file, loads it into an Aspose.Cells Workbook, iterates through the BuiltInDocumentProperties collection, and writes each property's Name and Value to the console while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all built‑in document properties
            foreach (var prop in workbook.BuiltInDocumentProperties)
            {
                // Output the property name and its value
                Console.WriteLine($"{prop.Name}: {prop.Value}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
