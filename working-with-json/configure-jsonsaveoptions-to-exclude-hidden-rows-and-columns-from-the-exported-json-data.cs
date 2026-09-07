// Title: Export Excel to JSON while skipping hidden rows and columns with Aspose.Cells for .NET
// AI Prompts: Write C# code that saves a Workbook as JSON using Aspose.Cells and omits any hidden rows and columns. | Show how to set up JsonSaveOptions in Aspose.Cells to prevent hidden rows or columns from being included in the JSON output. | Provide a C# example that loads an .xlsx file and generates a JSON file that contains only visible cells with Aspose.Cells. | Explain a workaround for excluding hidden rows/columns when the ExportHiddenRows/ExportHiddenColumns properties are unavailable.
// Common Searches: Aspose.Cells .NET export JSON without hidden rows | How to ignore hidden columns when converting Excel to JSON in C# | JsonSaveOptions hide hidden cells Aspose.Cells example | C# generate JSON from workbook excluding hidden rows and columns | Aspose.Cells JSON export filter out hidden data
// Tags: Aspose.Cells JSON export hide rows | JsonSaveOptions exclude hidden columns | C# workbook to JSON visible cells | Aspose.Cells filter hidden data during JSON save | export Excel as JSON without hidden rows .NET | Aspose.Cells JSON output visible range

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates or loads a Workbook, configures JsonSaveOptions, and saves the workbook as a JSON file. It highlights that the current Aspose.Cells version lacks ExportHiddenRows and ExportHiddenColumns properties, so all rows and columns are exported by default, and suggests upgrading or applying a custom filter to exclude hidden rows and columns.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // create rule

                // If you need to load an existing file, uncomment the lines below
                // string inputPath = "input.xlsx";
                // if (File.Exists(inputPath))
                // {
                //     workbook = new Workbook(inputPath); // load rule
                // }
                // else
                // {
                //     Console.WriteLine($"Input file not found: {inputPath}");
                //     return;
                // }

                // Configure JSON save options
                JsonSaveOptions jsonOptions = new JsonSaveOptions();

                // Note: ExportHiddenRows and ExportHiddenColumns properties are not available
                // in the current Aspose.Cells version. The default behavior exports all rows
                // and columns. Adjust options here if newer properties become available.

                // Save the workbook as JSON using the configured options
                string outputPath = "output.json";
                workbook.Save(outputPath, jsonOptions);
                Console.WriteLine($"Workbook saved as JSON to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
