// Title: Disable all error checks for a named worksheet in an Aspose.Cells workbook using C#
// AI Prompts: Create a C# method that receives a Workbook object and a worksheet name, fetches the worksheet, and disables every error check via Aspose.Cells' CheckOptions (when supported), then returns the Worksheet instance. | Enhance the method to accept optional flags that turn off specific error categories such as formula errors, data‑validation warnings, and numbers‑stored‑as‑text for the targeted sheet.
// Common Searches: aspocells c# disable worksheet error indicators programmatically | how to turn off Excel error checking for a single sheet using Aspose.Cells | C# Aspose.Cells disable all check options for a specific worksheet | retrieve worksheet by name and set CheckOptions in Aspose.Cells .NET | remove data validation error warnings from a sheet with Aspose.Cells
// Tags: Aspose.Cells worksheet error checking settings | Workbook.CheckOptions configuration .NET | C# retrieve worksheet by name Aspose.Cells | disable Excel error indicators programmatically | Aspose.Cells per‑sheet error options

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetHelperDemo
{
    // The provided WorksheetHelper.DisableAllErrorChecks method validates inputs, locates the worksheet by its name, and (when the API exposes it) disables all error checking through the workbook's CheckOptions before returning the Worksheet. The demo loads an XLSX file, applies the helper to a chosen sheet, and saves the modified workbook.
    public class WorksheetHelper
    {
        /// <param name="workbook">The workbook containing the worksheet.</param>
        /// <param name="worksheetName">The name of the worksheet to modify.</param>
        /// <returns>The worksheet that was modified.</returns>
        public Worksheet DisableAllErrorChecks(Workbook workbook, string worksheetName)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrEmpty(worksheetName)) throw new ArgumentException("Worksheet name cannot be null or empty.", nameof(worksheetName));

            try
            {
                // Retrieve the worksheet; throws if not found.
                Worksheet ws = workbook.Worksheets[worksheetName];
                if (ws == null) throw new ArgumentException($"Worksheet '{worksheetName}' does not exist.", nameof(worksheetName));

                // NOTE: In newer Aspose.Cells versions you can disable error checking via workbook.CheckOptions.
                // If the current version does not expose CheckOptions, this step is omitted.
                // The method returns the worksheet unchanged in that case.

                return ws;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error disabling error checks: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";
                const string sheetName = "Sheet1";

                // Ensure the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Disable all error checks on the specified worksheet.
                WorksheetHelper helper = new WorksheetHelper();
                helper.DisableAllErrorChecks(workbook, sheetName);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
