// Title: C# – Hide or show worksheets by name pattern with Aspose.Cells
// Description: Loads an Excel workbook, loops through every worksheet, hides those whose names start with a given prefix (e.g., "Hidden") by setting IsVisible = false, keeps the rest visible, and saves the modified file to a new location.
// Keywords: Aspose.Cells C# hide worksheet | worksheet visibility .NET | conditional sheet hide Aspose | Workbook.Save C# | IsVisible property Aspose.Cells | toggle Excel sheet visibility
// Common Searches: Aspose.Cells hide worksheets by prefix | C# set worksheet IsVisible property | How to hide multiple sheets with Aspose.Cells | Save workbook after changing sheet visibility | Toggle Excel sheet visibility programmatically
// Developer Intent: Programmatically adjust worksheet visibility according to naming rules and write the updated workbook to disk.
// Use Cases: Prepare a distribution package by automatically concealing internal tabs whose names begin with "Hidden" before sending the file to clients. | Enforce reporting standards by ensuring every sheet is explicitly marked visible or hidden based on business logic, then export the workbook.
// AI Prompts: Write C# code using Aspose.Cells that hides all worksheets containing the word "Temp" in their title and saves the workbook to a user‑specified path. | Show an example that iterates over a workbook's worksheets, applies a custom predicate to set the IsVisible flag, and writes the result to a new Excel file.

using System;
using Aspose.Cells;

// Loads an Excel workbook, loops through every worksheet, hides those whose names start with a given prefix (e.g., "Hidden") by setting IsVisible = false, keeps the rest visible, and saves the modified file to a new location.
class ToggleWorksheetVisibility
{
    static void Main()
    {
        // Load the workbook from an existing file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // uses Workbook(string) constructor

        // Loop through each worksheet and set visibility based on a condition
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Example condition: hide worksheets whose name starts with "Hidden"
            if (sheet.Name.StartsWith("Hidden", StringComparison.OrdinalIgnoreCase))
            {
                sheet.IsVisible = false; // hide the sheet
            }
            else
            {
                sheet.IsVisible = true; // ensure the sheet is visible
            }
        }

        // Save the modified workbook to a new file
        string outputPath = "output.xlsx";
        workbook.Save(outputPath); // uses Workbook.Save(string) method
    }
}
