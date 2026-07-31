// Title: Correct corrupted named ranges after worksheet rename with Aspose.Cells for .NET
// Description: Loads an Excel workbook, renames a worksheet, scans the workbook's NameCollection for defined names whose RefersTo formulas still contain the original sheet name, replaces the old sheet token with the new name, validates each corrected range via Name.GetRange(), and saves the repaired file.
// Keywords: Aspose.Cells | .NET | C# | named range | worksheet rename | RefersTo correction | NameCollection | repair broken named ranges | update defined names | Excel automation
// Common Searches: Aspose.Cells update named range after sheet rename | fix broken RefersTo references C# | detect corrupted named ranges in Excel | rename worksheet and correct defined names Aspose | how to repair named range links after sheet rename
// Developer Intent: Update any defined names that still reference the original worksheet after it has been renamed.
// Use Cases: Iterate through all workbook names to locate RefersTo strings containing the old sheet name. | Replace the outdated sheet token with the new worksheet name and assign the corrected RefersTo value. | Validate each updated name by retrieving its Range object to ensure the reference is functional. | Save the workbook to produce an Excel file free of broken named‑range links.
// AI Prompts: Generate C# code using Aspose.Cells that scans a workbook's NameCollection, replaces old worksheet names in RefersTo strings with a new name, and verifies each range. | Provide a method to detect and fix corrupted named ranges after a sheet rename, handling missing RefersTo values and exceptions, then save the corrected workbook. | Explain best practices for maintaining named ranges when renaming worksheets in Aspose.Cells, including validation and error handling.

using System;
using System.IO;
using Aspose.Cells;

namespace NamedRangeCorrectionDemo
{
    // Loads an Excel workbook, renames a worksheet, scans the workbook's NameCollection for defined names whose RefersTo formulas still contain the original sheet name, replaces the old sheet token with the new name, validates each corrected range via Name.GetRange(), and saves the repaired file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_corrected.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the existing workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Original sheet name before renaming
            const string oldSheetName = "Sheet1";

            // Rename the first worksheet
            const string newSheetName = "RenamedSheet";
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = newSheetName;

            // Iterate through all defined names in the workbook
            NameCollection names = workbook.Worksheets.Names;
            foreach (Name name in names)
            {
                string refersTo = name.RefersTo;
                if (string.IsNullOrEmpty(refersTo))
                    continue; // Skip names that do not refer to a range

                // Detect references that still contain the old sheet name
                string oldSheetToken = oldSheetName + "!";
                if (refersTo.Contains(oldSheetToken))
                {
                    // Update the reference to use the new sheet name
                    string correctedRefersTo = refersTo.Replace(oldSheetToken, newSheetName + "!");
                    name.RefersTo = correctedRefersTo;

                    // Verify that the corrected range can be retrieved
                    try
                    {
                        // Use fully qualified Aspose.Cells.Range to avoid ambiguity with System.Range
                        Aspose.Cells.Range correctedRange = name.GetRange();
                        Console.WriteLine($"Name '{name.Text}' corrected to range {correctedRange.Address}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to retrieve range for name '{name.Text}': {ex.Message}");
                    }
                }
            }

            // Save the corrected workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
