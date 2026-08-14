// Title: C# – Hide or Show Worksheets by Name Prefix with Aspose.Cells
// Description: Loads an Excel workbook, iterates all worksheets, hides those whose names begin with a given prefix (e.g., "Temp"), ensures the rest are visible, and saves the modified file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | C# | worksheet visibility | hide worksheet | show worksheet | IsVisible property | load workbook | save workbook | conditional sheet hide | Excel automation
// Common Searches: Aspose.Cells hide worksheet C# | set worksheet visibility Aspose.Cells .NET | toggle worksheet visibility based on name | load and save workbook after changing sheet visibility | C# code to hide sheets starting with Temp
// Developer Intent: Programmatically hide or reveal worksheets according to a naming rule and write the changes back to the file.
// Use Cases: Clean up generated reports by hiding temporary tabs before distribution. | Create a user‑friendly template that displays only the required worksheets. | Enforce naming conventions in automated workflows by concealing placeholder sheets.
// AI Prompts: Write C# code with Aspose.Cells that hides all worksheets whose name contains "Draft" and saves the result as a new file. | Provide an example that toggles worksheet visibility using a custom predicate and logs each sheet that was hidden or shown. | Create a function that accepts a name prefix, sets IsVisible = false for matching sheets, returns the count of hidden sheets, and saves the workbook.

using System;
using Aspose.Cells;

namespace WorksheetVisibilityToggle
{
    // Loads an Excel workbook, iterates all worksheets, hides those whose names begin with a given prefix (e.g., "Temp"), ensures the rest are visible, and saves the modified file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook
            string inputPath = "input.xlsx";

            // Load the workbook from file (uses the Workbook(string) constructor)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and toggle visibility based on a condition
            // Example condition: hide worksheets whose name starts with "Temp"
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.StartsWith("Temp", StringComparison.OrdinalIgnoreCase))
                {
                    // Hide the worksheet
                    sheet.IsVisible = false;
                }
                else
                {
                    // Ensure the worksheet is visible
                    sheet.IsVisible = true;
                }
            }

            // Save the updated workbook (uses the Workbook.Save(string) method)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}' with updated worksheet visibility.");
        }
    }
}
