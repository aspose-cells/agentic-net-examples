// Title: C# – Iterate Aspose.Cells Workbook.LoadWarnings and Print Each Warning
// Description: Loads an Excel file with Aspose.Cells, accesses the Workbook.LoadWarnings collection (using reflection when the property is unavailable), enumerates WarningInfo objects, extracts their Description, and writes each warning to the console while handling missing‑property scenarios and runtime exceptions.
// Keywords: Aspose.Cells LoadWarnings C# | Workbook.LoadWarnings enumeration | display Excel load warnings .NET | reflection access LoadWarnings property | Aspose.Cells warning messages | C# Excel workbook warnings
// Common Searches: how to read load warnings from Aspose.Cells workbook | iterate Workbook.LoadWarnings C# example | access LoadWarnings property via reflection Aspose.Cells | list warning descriptions after loading Excel with Aspose.Cells | Aspose.Cells missing LoadWarnings property workaround
// Developer Intent: Enumerate and output all load warning messages generated when a workbook is opened with Aspose.Cells.
// Use Cases: Log unsupported features or data‑loss warnings after opening user‑provided Excel files. | Show end‑users a summary of load warnings so they can correct source files before further processing. | Trigger conditional logic (e.g., abort or modify processing) based on critical load warnings.
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells and prints each load warning description without using reflection. | Create a reusable method that returns a List<string> of warning messages from a Workbook, handling cases where the LoadWarnings property is absent. | Write error‑handling logic that captures load warnings, writes them to a log file, and still allows further workbook operations.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadWarningsDemo
{
    // Loads an Excel file with Aspose.Cells, accesses the Workbook.LoadWarnings collection (using reflection when the property is unavailable), enumerates WarningInfo objects, extracts their Description, and writes each warning to the console while handling missing‑property scenarios and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file '{filePath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook using the standard constructor
                Workbook workbook = new Workbook(filePath);

                // Attempt to retrieve load warnings via reflection.
                // Some older Aspose.Cells versions may not expose the LoadWarnings property.
                var loadWarningsProp = typeof(Workbook).GetProperty("LoadWarnings");
                if (loadWarningsProp != null)
                {
                    var warnings = loadWarningsProp.GetValue(workbook) as System.Collections.IEnumerable;
                    if (warnings != null)
                    {
                        foreach (var warningObj in warnings)
                        {
                            // Each warning object is of type WarningInfo; retrieve its Description property.
                            var descriptionProp = warningObj.GetType().GetProperty("Description");
                            if (descriptionProp != null)
                            {
                                string description = descriptionProp.GetValue(warningObj) as string;
                                Console.WriteLine($"Warning: {description}");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("LoadWarnings property is not available in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            // Optional: keep console window open when running outside an IDE
            Console.WriteLine("Processing completed.");
        }
    }
}
