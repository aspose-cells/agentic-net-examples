// Title: Read and log all built‑in document properties from an Excel workbook using Aspose.Cells for .NET
// Common Searches: c# read built-in document properties from Excel using Aspose.Cells | list all workbook built-in properties Aspose.Cells example | iterate over BuiltInDocumentProperties collection in Aspose.Cells | log Excel file metadata with Aspose.Cells C# console app | handle missing Excel file when reading properties Aspose.Cells
// Tags: Aspose.Cells BuiltInDocumentProperties enumeration | C# iterate Excel workbook metadata collection | validate input.xlsx existence before Aspose.Cells load | console logging of Excel document properties in .NET | Aspose.Cells workbook loading error handling

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that input.xlsx exists, loads it with Aspose.Cells, accesses the BuiltInDocumentProperties collection, iterates through each property, and writes the property name and value to the console while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the Excel workbook
            var workbook = new Workbook(inputPath);

            // Get the collection of built‑in document properties
            var builtInProps = workbook.BuiltInDocumentProperties;

            // Iterate through each property and display its name and value
            foreach (var prop in builtInProps)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
