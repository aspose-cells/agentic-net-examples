// Title: C# – Load a massive workbook with LoadOptions.MemorySetting = MemoryPreference to prevent OutOfMemoryException (Aspose.Cells)
// Description: Demonstrates how to enable MemoryPreference in Aspose.Cells LoadOptions for .NET, create a placeholder workbook if the source file is missing, load a large Excel file with reduced memory pressure, display worksheet count, and save the result.
// Keywords: Aspose.Cells LoadOptions MemoryPreference | MemorySetting MemoryPreference C# | large workbook OutOfMemoryException Aspose | load massive Excel file .NET | Aspose.Cells memory management
// Common Searches: Aspose.Cells MemoryPreference example | LoadOptions MemorySetting large workbook | prevent OutOfMemoryException Aspose.Cells | C# load big Excel file with low memory usage | Aspose.Cells placeholder workbook if file missing
// Developer Intent: Enable MemoryPreference in LoadOptions to open a huge Excel workbook without exhausting system memory.
// Use Cases: Processing multi‑gigabyte Excel files on a server with limited RAM. | Reading large workbooks for analytics while keeping the application responsive. | Automatically creating a placeholder workbook when the target file does not exist, then loading it in memory‑preference mode.
// AI Prompts: Write C# code that loads an Excel file using Aspose.Cells with LoadOptions.MemorySetting set to MemoryPreference and handles a missing file gracefully. | Show how to configure LoadOptions for MemoryPreference and iterate through worksheets safely in Aspose.Cells. | Explain the impact of MemoryPreference on memory consumption and when to apply it in Aspose.Cells projects.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to enable MemoryPreference in Aspose.Cells LoadOptions for .NET, create a placeholder workbook if the source file is missing, load a large Excel file with reduced memory pressure, display worksheet count, and save the result.
class Program
{
    static void Main()
    {
        string inputFile = "massive_workbook.xlsx";
        string outputFile = "massive_workbook_processed.xlsx";

        try
        {
            // Ensure the input file exists; create a placeholder if it does not.
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found. Creating an empty workbook as a placeholder.");
                var placeholder = new Workbook();
                placeholder.Save(inputFile, SaveFormat.Xlsx);
            }

            // Enable memory‑preference mode for large workbooks.
            var loadOptions = new LoadOptions
            {
                MemorySetting = MemorySetting.MemoryPreference
            };

            // Load the workbook with the specified options.
            var workbook = new Workbook(inputFile, loadOptions);

            // Example operation: display the number of worksheets loaded.
            Console.WriteLine($"Worksheets loaded: {workbook.Worksheets.Count}");

            // Save the processed workbook.
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
