// Title: Read the Author built‑in document property from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, checks if the file exists, and returns the value of the built‑in Author property. | Show a complete example that loads a workbook, accesses BuiltInDocumentProperties, extracts the creator name, and includes exception handling. | Provide a snippet that prints the Author metadata of an Excel file using Aspose.Cells, handling missing‑file scenarios.
// Common Searches: aspnet read author property from excel using Aspose.Cells | C# Aspose.Cells get built‑in document properties Author | how to retrieve creator name from .xlsx with Aspose.Cells library | example code for checking file existence before reading Excel metadata Aspose.Cells | Aspose.Cells built‑in properties read Author value C#
// Tags: Aspose.Cells read built-in document properties | C# retrieve Excel Author metadata | load workbook and access BuiltInDocumentProperties | handle missing file with Aspose.Cells | extract creator name from .xlsx

using System;
using System.IO;
using Aspose.Cells;

// // This C# program verifies that input.xlsx exists, loads it with Aspose.Cells, accesses the workbook's BuiltInDocumentProperties, reads the "Author" property (the creator), prints the author name, and catches any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Access built‑in document properties
            var builtInProps = workbook.BuiltInDocumentProperties;

            // Retrieve the Author property (creator of the file)
            string author = builtInProps["Author"]?.ToString();

            // Display the author name
            Console.WriteLine("Author: " + author);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
