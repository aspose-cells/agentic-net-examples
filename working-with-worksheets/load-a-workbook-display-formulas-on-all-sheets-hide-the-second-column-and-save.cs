// Title: Show Formulas & Hide Column B on Every Worksheet Using Aspose.Cells for .NET
// Description: Load an Excel workbook, set each worksheet to display formulas (ShowFormulas = true), hide the second column (B) across all sheets, and save the modified file with Aspose.Cells in C#.
// Keywords: Aspose.Cells C# | show formulas Excel | hide column B | iterate worksheets | modify Excel workbook | save workbook Aspose | batch worksheet settings | Excel column hide .NET | display formulas Aspose.Cells | Aspose.Cells .NET example
// Common Searches: Aspose.Cells hide column in all worksheets | Show formulas in Excel using Aspose.Cells .NET | Set ShowFormulas for every sheet programmatically | Batch hide column B with Aspose.Cells | Modify multiple worksheets Aspose.Cells C#
// Developer Intent: Load a workbook, make formulas visible on every sheet, hide column B, and save the updated file.
// Use Cases: Create an audit‑ready copy where all formulas are visible while confidential data in column B is concealed. | Prepare a distribution version that shows calculation logic for transparency but hides internal identifiers. | Apply consistent view settings before exporting the workbook to PDF or other formats.
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, sets ShowFormulas = true for each worksheet, hides column index 1, and saves the file. | Explain how to iterate over all worksheets in an Excel file using Aspose.Cells to display formulas and hide a specific column. | Provide troubleshooting steps if hiding column B does not appear in the saved workbook when using Aspose.Cells for .NET.

using Aspose.Cells;
using System;

// Load an Excel workbook, set each worksheet to display formulas (ShowFormulas = true), hide the second column (B) across all sheets, and save the modified file with Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Apply settings to each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Show formulas instead of their results
            sheet.ShowFormulas = true;

            // Hide the second column (zero‑based index 1)
            sheet.Cells.HideColumn(1);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
