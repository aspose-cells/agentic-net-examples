// Title: Show Formulas on Every Worksheet and Hide Column B Using Aspose.Cells for .NET (C#)
// Description: C# sample that opens an Excel workbook with Aspose.Cells, enables formula display on each sheet, hides column B (index 1) throughout the file, and writes the changes to a new workbook.
// Keywords: Aspose.Cells C# show formulas | hide column B Aspose.Cells | iterate worksheets Aspose.Cells | save modified workbook Aspose.Cells | Excel formula view .NET | column visibility Aspose.Cells | bulk worksheet operations Aspose.Cells | Excel auditing C#
// Common Searches: Aspose.Cells display formulas in all sheets | How to hide a column in every worksheet using Aspose.Cells | Set ShowFormulas property for a workbook in C# | C# hide column B across an Excel file with Aspose.Cells | Export workbook with formulas visible Aspose.Cells
// Developer Intent: Load an existing workbook, make formulas visible on all sheets, conceal column B, and save the updated file.
// Use Cases: Auditing: produce a copy of a spreadsheet where formulas are shown for review while sensitive data in column B is hidden. | Debugging: generate a version of a workbook that reveals all calculations and removes column B to simplify troubleshooting. | Documentation: create a printable Excel file that lists formulas for reference and omits column B to reduce visual clutter.
// AI Prompts: Generate C# code with Aspose.Cells that toggles ShowFormulas for every worksheet and hides column C before saving. | Provide an example that loads an Excel file, displays formulas, hides a list of specified columns on all sheets, and outputs a new workbook. | Explain how to programmatically control formula visibility and column hiding based on user input using Aspose.Cells in a .NET application.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# sample that opens an Excel workbook with Aspose.Cells, enables formula display on each sheet, hides column B (index 1) throughout the file, and writes the changes to a new workbook.
    class ShowFormulasAndHideColumn
    {
        static void Main()
        {
            // Load an existing workbook from file
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Show formulas instead of calculated values
                sheet.ShowFormulas = true;

                // Hide the second column (index 1, column B)
                sheet.Cells.HideColumn(1);
            }

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}
