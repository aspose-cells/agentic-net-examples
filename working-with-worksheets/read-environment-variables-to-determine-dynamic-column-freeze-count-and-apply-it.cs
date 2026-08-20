// Title: Dynamic column freeze in Aspose.Cells via environment variable (C#)
// Description: Creates a new Workbook, reads the FREEZE_COLUMNS environment variable, validates it as a non‑negative integer, and freezes that many columns (no rows) on every worksheet using FreezePanes. If the variable is missing, zero, or invalid, panes are left unfrozen before the file is saved.
// Keywords: Aspose.Cells | FreezePanes | C# | .NET | environment variable | dynamic column freeze | unfreeze panes | Excel automation | read env variable C#
// Common Searches: Aspose.Cells freeze columns from environment variable | C# set FreezePanes using env var | How to unfreeze panes in Aspose.Cells | Apply FreezePanes to all worksheets Aspose.Cells | Dynamic column freeze Aspose.Cells .NET
// Developer Intent: Read an environment variable and use its value to freeze that many columns in every worksheet of an Aspose.Cells workbook.
// Use Cases: Configure column freezing per deployment environment without code changes | Disable freezing by leaving FREEZE_COLUMNS unset or setting it to 0 | Apply identical freeze settings across all worksheets in a newly created workbook | Toggle freeze behavior in CI/CD pipelines via environment configuration
// AI Prompts: Generate C# code that reads an environment variable named FREEZE_COLUMNS and applies FreezePanes to each worksheet in an Aspose.Cells workbook, handling invalid values gracefully. | Show how to extend the example to also freeze rows based on a separate environment variable. | Explain how to unit‑test the dynamic column‑freeze logic with Aspose.Cells in a .NET test project. | Provide a PowerShell script that sets the FREEZE_COLUMNS variable before running the C# application.

using System;
using Aspose.Cells;

namespace FreezeColumnsFromEnv
{
    // Creates a new Workbook, reads the FREEZE_COLUMNS environment variable, validates it as a non‑negative integer, and freezes that many columns (no rows) on every worksheet using FreezePanes. If the variable is missing, zero, or invalid, panes are left unfrozen before the file is saved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Read the environment variable that specifies how many columns to freeze.
            // If the variable is not set or is invalid, default to 0 (no freezing).
            string envValue = Environment.GetEnvironmentVariable("FREEZE_COLUMNS");
            int freezeColumns = 0;
            if (!string.IsNullOrEmpty(envValue) && int.TryParse(envValue, out int parsed))
            {
                // Ensure the value is non‑negative.
                freezeColumns = Math.Max(0, parsed);
            }

            // Apply the freeze setting to each worksheet in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (freezeColumns > 0)
                {
                    // Freeze panes at row index 0 and the specified column index.
                    // freezedRows = 0 (no frozen rows), freezedColumns = freezeColumns.
                    sheet.FreezePanes(0, freezeColumns, 0, freezeColumns);
                }
                else
                {
                    // Ensure panes are not frozen when the count is zero.
                    sheet.UnFreezePanes();
                }
            }

            // Save the workbook to a file.
            workbook.Save("FreezeColumnsResult.xlsx");
        }
    }
}
