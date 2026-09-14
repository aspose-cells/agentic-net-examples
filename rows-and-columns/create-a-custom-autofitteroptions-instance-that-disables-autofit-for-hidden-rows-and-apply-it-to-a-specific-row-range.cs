// Title: How to disable auto‑fit for hidden rows using AutoFitterOptions and apply AutoFitRows to a specific range in Aspose.Cells for .NET
// AI Prompts: Write C# that creates an AutoFitterOptions instance with IgnoreHidden = true and calls Worksheet.AutoFitRows(startRow, endRow, options) so only visible rows are resized. | Provide a code example that hides selected rows, applies a custom auto‑fit options object, autofits a defined row range, and saves the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# AutoFitRows ignore hidden rows in a range | prevent hidden rows from affecting row height auto‑fit in .NET Excel | configure AutoFitterOptions to skip hidden rows for specific rows | apply autofit to rows 0‑5 while excluding hidden rows using Aspose.Cells | custom auto‑fit options example for row height adjustment in C#
// Tags: AutoFitterOptions.IgnoreHidden property | AutoFitRows with custom options Aspose.Cells | hide rows before autofit .NET | row range autofit Aspose.Cells | skip hidden rows Excel autofit C#

using System;
using Aspose.Cells;

// The sample creates a workbook, fills six rows with long text, hides rows 2 and 4, sets AutoFitterOptions.IgnoreHidden to true, autofits rows 0‑5 using those options, and saves the file as AutoFitRowsIgnoreHidden.xlsx.
class AutoFitHiddenRowsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to several rows
        for (int i = 0; i < 6; i++)
        {
            sheet.Cells[i, 0].PutValue($"Row {i} - this is a long text that may require autofit");
        }

        // Hide some rows (row indices are zero‑based)
        sheet.Cells.Rows[2].IsHidden = true;
        sheet.Cells.Rows[4].IsHidden = true;

        // Create AutoFitterOptions that ignores hidden rows during autofit
        AutoFitterOptions options = new AutoFitterOptions
        {
            IgnoreHidden = true   // hidden rows will not be considered for autofit
        };

        // Apply autofit to rows 0 through 5 using the custom options
        sheet.AutoFitRows(0, 5, options);

        // Save the workbook
        workbook.Save("AutoFitRowsIgnoreHidden.xlsx");
    }
}
