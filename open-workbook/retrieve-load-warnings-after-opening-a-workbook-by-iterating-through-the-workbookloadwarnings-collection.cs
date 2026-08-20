// Title: C# – Retrieve Workbook LoadWarnings with Aspose.Cells after Opening an Excel File
// Description: Loads an Excel workbook using Aspose.Cells, accesses the Workbook.LoadWarnings collection (using reflection for older versions), iterates through each warning, prints its Type and Description, and saves the workbook. Includes error handling for missing files and unavailable properties.
// Keywords: Aspose.Cells LoadWarnings | Workbook.LoadWarnings C# | iterate load warnings | Excel load warnings .NET | reflection access LoadWarnings | Aspose.Cells compatibility diagnostics | C# load options Excel | Aspose.Cells version check | retrieve workbook warnings | load warnings collection
// Common Searches: how to get load warnings with Aspose.Cells C# | Workbook.LoadWarnings iteration example | access LoadWarnings property via reflection | display warning type and description Aspose.Cells | Aspose.Cells load warnings not available in older version
// Developer Intent: Extract and show any load warnings produced when opening an Excel workbook with Aspose.Cells.
// Use Cases: Log warnings to identify unsupported features in user‑uploaded spreadsheets. | Validate workbook integrity before data processing by checking for load warnings. | Provide end‑user feedback about compatibility issues detected during file import.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells and prints all load warnings without using reflection. | Show how to filter Workbook.LoadWarnings by warning type after loading a workbook. | Explain strategies for handling the absence of the LoadWarnings property in older Aspose.Cells releases.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadWarningsDemo
{
    // Loads an Excel workbook using Aspose.Cells, accesses the Workbook.LoadWarnings collection (using reflection for older versions), iterates through each warning, prints its Type and Description, and saves the workbook. Includes error handling for missing files and unavailable properties.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be loaded
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Create LoadOptions for the desired format (no LoadWarnings property needed)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Load the workbook using the LoadOptions
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Try to retrieve the LoadWarnings collection via reflection (covers versions where the property may be missing)
                var warningsProp = typeof(Workbook).GetProperty("LoadWarnings");
                if (warningsProp != null)
                {
                    var warnings = warningsProp.GetValue(workbook) as IEnumerable;
                    if (warnings != null)
                    {
                        foreach (var warningObj in warnings)
                        {
                            // Use reflection to read WarningInfo members
                            var typeProp = warningObj.GetType().GetProperty("Type");
                            var descProp = warningObj.GetType().GetProperty("Description");

                            var typeValue = typeProp?.GetValue(warningObj);
                            var descValue = descProp?.GetValue(warningObj);

                            Console.WriteLine($"Warning Type: {typeValue}");
                            Console.WriteLine($"Description : {descValue}");
                            Console.WriteLine();
                        }
                    }
                }
                else
                {
                    Console.WriteLine("LoadWarnings property is not available in this version of Aspose.Cells.");
                }

                // Optionally, save the workbook to verify it is still functional
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
