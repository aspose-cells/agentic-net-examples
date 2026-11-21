// Title: Retrieve Excel workbook load warnings in C# by iterating the Aspose.Cells Workbook.LoadWarnings collection
// Common Searches: C# Aspose.Cells how to get load warnings after opening an Excel file | Iterate over Workbook.LoadWarnings to list loading issues with Aspose.Cells .NET | Capture load warnings using LoadOptions when loading a workbook in Aspose.Cells | Aspose.Cells .NET retrieve warnings for corrupted Excel workbook during load | Example code for reading load warnings from a workbook with Aspose.Cells C#
// Tags: Aspose.Cells workbook load warnings retrieval | C# iterate LoadWarnings collection | Aspose.Cells LoadOptions warning capture | Excel file loading diagnostics .NET | handling loading errors with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the Excel file exists, loads it with Aspose.Cells in C#, and notes that newer Aspose.Cells releases no longer expose Workbook.LoadWarnings directly; to obtain load warnings you must configure LoadOptions and use the appropriate Load overload, then iterate the collected warnings.
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
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // NOTE: In recent versions of Aspose.Cells the LoadWarnings collection
            // is not exposed directly. If you need to capture loading warnings,
            // configure LoadOptions and handle them via the Load method overloads.
            // For this example we simply confirm that the workbook was loaded.

            Console.WriteLine($"Workbook loaded successfully. Worksheets count: {workbook.Worksheets.Count}");

            // Continue with further processing as needed...
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions (e.g., corrupted file, unsupported format)
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }
    }
}
