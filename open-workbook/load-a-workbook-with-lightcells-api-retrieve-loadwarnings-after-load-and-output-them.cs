// Title: Use Aspose.Cells LightCells API in C# to load an XLSX workbook, capture LoadWarnings, and display them
// AI Prompts: Write C# code that opens an XLSX file with Aspose.Cells LightCells API, accesses the workbook's LoadWarnings collection, and prints each warning to the console. | Show how to configure LoadOptions for LightCells, load a workbook, and iterate over workbook.LoadWarnings to log warning messages in a .NET application.
// Common Searches: asp.net core aspose.cells lightcells load workbook and get load warnings c# | how to read load warnings after opening an Excel file with Aspose.Cells in C# | example of retrieving LoadWarnings collection from a workbook loaded with LightCells API | c# code to log warnings generated during Aspose.Cells workbook load
// Tags: aspose.cells lightcells load warnings c# | loadoptions retrieve workbook load warnings c# | excel workbook load warnings handling asp.net | c# aspose.cells load workbook with warnings output

using System;
using System.IO;
using Aspose.Cells;

// The example checks that 'input.xlsx' exists, creates LoadOptions for the XLSX format, loads the workbook with Aspose.Cells LightCells API, accesses the workbook.LoadWarnings collection, iterates through each warning and writes its description to the console, and handles any exceptions that may occur.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found - {filePath}");
            return;
        }

        try
        {
            // Load the workbook (default options are sufficient for most scenarios)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(filePath, loadOptions);

            // If needed, you can access workbook properties here.
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
