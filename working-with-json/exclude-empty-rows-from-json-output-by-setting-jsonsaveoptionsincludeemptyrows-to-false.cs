// Title: How to export an Excel workbook to JSON without empty rows using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures JsonSaveOptions to omit empty rows, and saves the result as a .json file. | Provide a version‑safe snippet that sets the ExportEmptyRows property on JsonSaveOptions only when the property exists at runtime. | Show how to add basic error handling for a missing input workbook while performing the Excel‑to‑JSON conversion with Aspose.Cells.
// Common Searches: Aspose.Cells C# export Excel to JSON exclude blank rows | Set ExportEmptyRows false in JsonSaveOptions for .NET 6+ | Prevent empty rows when saving a workbook as JSON using Aspose.Cells | Conditional check for ExportEmptyRows property in older Aspose.Cells SDKs
// Tags: aspose.cells jsonsaveoptions exportemptyrows false | c# workbook to json conversion aspose.cells | reflection based property setting aspnet | conditional compilation for sdk compatibility | json export without empty rows aspose.cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    // The example loads an Excel workbook, creates a JsonSaveOptions object, uses reflection (under .NET 6 or later) to set ExportEmptyRows to false when the property is available, and saves the workbook as a JSON file, thereby omitting any empty rows from the generated JSON.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = "input.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Set JSON save options (exclude empty rows if supported)
                JsonSaveOptions jsonOptions = new JsonSaveOptions();

                // Some versions expose ExportEmptyRows; set it to false when available
                // This conditional compilation avoids errors on older SDKs
#if NET6_0_OR_GREATER
                if (jsonOptions.GetType().GetProperty("ExportEmptyRows") != null)
                {
                    jsonOptions.GetType().GetProperty("ExportEmptyRows")?.SetValue(jsonOptions, false);
                }
#endif

                string outputFile = "output.json";

                // Save the workbook as JSON
                workbook.Save(outputFile, jsonOptions);
                Console.WriteLine($"Workbook successfully saved to JSON: {outputFile}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
