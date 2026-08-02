// Title: C# – Enumerate Aspose.Cells Workbook LoadWarnings and Print Their Descriptions
// Description: Loads an Excel file with Aspose.Cells, detects the Workbook.LoadWarnings property (using reflection for version‑agnostic support), iterates through each WarningInfo entry, and writes the warning text to the console while gracefully handling missing members or runtime errors.
// Keywords: Aspose.Cells LoadWarnings C# | Workbook warning enumeration .NET | print Excel load warnings | reflection access LoadWarnings | Aspose.Cells version compatibility | log workbook warnings
// Common Searches: how to read load warnings from Aspose.Cells workbook | C# example for Workbook.LoadWarnings iteration | display Aspose.Cells warning messages in console | check if LoadWarnings property exists in Aspose.Cells | retrieve warning descriptions after opening Excel file
// Developer Intent: The developer needs to list every warning generated while loading an Excel workbook with Aspose.Cells and output each warning's description to the console.
// Use Cases: Create a diagnostic log of all load‑time warnings for user‑uploaded spreadsheets. | Show warning details in a UI panel so users can correct problematic data before further processing. | Trigger conditional logic (e.g., abort, fallback) when specific warnings indicate unsupported features.
// AI Prompts: Write C# code that accesses Workbook.LoadWarnings directly (no reflection) for the latest Aspose.Cells release and prints each warning description. | Show how to fallback to a custom warning collector when the LoadWarnings property is absent in older Aspose.Cells versions. | Generate a unit test that loads a workbook with known issues and asserts that all warning descriptions appear in the console output.

using System;
using System.Collections;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsLoadWarningsDemo
{
    // Loads an Excel file with Aspose.Cells, detects the Workbook.LoadWarnings property (using reflection for version‑agnostic support), iterates through each WarningInfo entry, and writes the warning text to the console while gracefully handling missing members or runtime errors.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook using the standard constructor (load rule)
                Workbook workbook = new Workbook(filePath);

                // Attempt to retrieve load warnings via reflection (compatible with multiple Aspose.Cells versions)
                PropertyInfo warningsProp = typeof(Workbook).GetProperty("LoadWarnings", BindingFlags.Public | BindingFlags.Instance);
                if (warningsProp != null)
                {
                    var warnings = warningsProp.GetValue(workbook) as IEnumerable;
                    if (warnings != null)
                    {
                        foreach (object warningObj in warnings)
                        {
                            if (warningObj is WarningInfo warning)
                            {
                                Console.WriteLine($"Warning: {warning.Description}");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("LoadWarnings property is not available in the current Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display an error message
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}
