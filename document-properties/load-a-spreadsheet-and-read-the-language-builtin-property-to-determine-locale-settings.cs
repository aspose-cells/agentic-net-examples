// Title: Read the Language built‑in document property from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read Language built‑in property from .xlsx file | how to get Excel document language setting with Aspose.Cells in .NET | retrieve built‑in Language metadata from workbook using Aspose.Cells | Aspose.Cells example to read locale information from Excel file | load workbook and access Language property via BuiltInDocumentProperties in C#
// Tags: read built-in document properties Aspose.Cells | retrieve language metadata Excel C# | Aspose.Cells workbook locale property | C# get Excel language built-in property | Aspose.Cells built-in properties access example

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies that 'input.xlsx' exists, loads it into an Aspose.Cells Workbook, accesses the BuiltInDocumentProperties collection, safely reads the 'Language' property, prints its value (or 'N/A' if absent), and handles any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string filePath = "input.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File \"{filePath}\" not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Access built‑in document properties
            var builtInProps = workbook.BuiltInDocumentProperties;

            // Retrieve the "Language" property safely using the indexer
            var languageProp = builtInProps["Language"];
            string languageValue = languageProp != null && languageProp.Value != null
                ? languageProp.Value.ToString()
                : "N/A";

            // Display the language setting
            Console.WriteLine("Document Language: " + languageValue);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
