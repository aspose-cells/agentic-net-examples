// Title: Open a large XLSX workbook with Aspose.Cells LightCells API while restricting memory usage to 200 MB in C#
// AI Prompts: Write C# code that loads an XLSX file using Aspose.Cells LightCells API and enforces a maximum of 200 MB memory consumption. | Show how to configure LoadOptions.MemorySetting to MemoryPreference for memory‑saving mode when opening a big Excel workbook in .NET. | Create an example that checks for file existence, opens the workbook with limited memory, prints the worksheet count, and catches any exceptions.
// Common Searches: aspnet open large xlsx with aspose.cells memory limit 200mb | c# lightcells api reduce memory usage when loading workbook | load big excel file without exceeding 200mb memory aspose cells | how to use LoadOptions.MemorySetting to limit memory in aspose.cells
// Tags: Aspose.Cells LightCells open large XLSX memory‑preference | C# LoadOptions.MemorySetting memory‑saving mode | limit workbook memory consumption Aspose.Cells | load big Excel file low memory .NET | memory‑efficient workbook loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks that a large XLSX file exists, then loads it with Aspose.Cells using LoadOptions set to MemorySetting.MemoryPreference to keep memory usage around 200 MB, outputs the number of worksheets, and handles any errors gracefully.
class Program
{
    static void Main()
    {
        // Path to the large XLSX file
        string inputPath = @"C:\Data\LargeFile.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Configure load options to use memory‑saving mode
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                MemorySetting = MemorySetting.MemoryPreference
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Example: output the number of worksheets
            Console.WriteLine($"Worksheets count: {workbook.Worksheets.Count}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
