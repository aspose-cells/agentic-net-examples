// Title: Aspose.Cells .NET – Use AutoFitterOptions to skip hidden rows when autofitting a row range
// Description: Demonstrates how to create a workbook, hide a specific row, configure an AutoFitterOptions object with IgnoreHidden set to true, and apply worksheet.AutoFitRows(start, end, options) to adjust the height of rows 0‑4 while leaving the hidden row unchanged, then save the file.
// Keywords: Aspose.Cells AutoFitterOptions | IgnoreHidden property | AutoFitRows range .NET | skip hidden rows Excel | C# autofit rows specific range
// Common Searches: Aspose.Cells how to ignore hidden rows during autofit | AutoFitRows with custom options in C# | disable autofit for hidden rows Aspose.Cells | fit only visible rows in a worksheet range | AutoFitterOptions example .NET
// Developer Intent: Create an AutoFitterOptions instance that excludes hidden rows and apply it to a defined row interval.
// Use Cases: Generating Excel reports where hidden rows must retain their original height while visible rows auto‑adjust. | Exporting data tables with mixed visibility, ensuring only displayed rows are resized. | Applying different autofit settings to multiple sections of a worksheet, such as ignoring hidden rows in one range and fitting all rows in another.
// AI Prompts: Write C# code using Aspose.Cells to autofit rows 5‑10 while ignoring any hidden rows in that range. | Show how to set AutoFitterOptions.IgnoreHidden to true and apply AutoFitRows to a worksheet that contains both hidden and visible rows. | Provide an example that toggles the IgnoreHidden flag based on a method parameter, fits a given row range, and saves the workbook.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, hide a specific row, configure an AutoFitterOptions object with IgnoreHidden set to true, and apply worksheet.AutoFitRows(start, end, options) to adjust the height of rows 0‑4 while leaving the hidden row unchanged, then save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add sample data to rows 0‑4
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue($"Row {i} contains a long piece of text that may require autofit.");
        }

        // Hide row 2 (zero‑based index)
        worksheet.Cells.Rows[2].IsHidden = true;

        // Create AutoFitterOptions that ignores hidden rows/columns
        AutoFitterOptions options = new AutoFitterOptions
        {
            IgnoreHidden = true,   // disables auto‑fit for hidden rows
            OnlyAuto = false       // optional: fit all rows, not only those without custom height
        };

        // Apply autofit to rows 0 through 4 using the custom options
        worksheet.AutoFitRows(0, 4, options);

        // Save the workbook
        workbook.Save("AutoFitRowsIgnoreHidden.xlsx");
    }
}
