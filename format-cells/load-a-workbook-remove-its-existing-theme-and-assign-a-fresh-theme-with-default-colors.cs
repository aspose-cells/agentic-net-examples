// Title: C# – Reset an Excel workbook’s theme to the default using Aspose.Cells
// Description: Loads an existing .xlsx file, creates a fresh workbook that contains Aspose.Cells’ built‑in default theme, copies that theme onto the loaded workbook with CopyTheme, and saves the result. The process removes any custom theme and restores the standard default colors.
// Keywords: Aspose.Cells | C# | .NET | reset Excel theme | default theme | CopyTheme | remove custom theme | workbook theme manipulation | Excel theme example | GitHub sample
// Common Searches: how to reset Excel theme with Aspose.Cells C# | copy default theme to existing workbook Aspose.Cells | remove custom theme from .xlsx programmatically | Aspose.Cells example for theme replacement | C# code to apply default Excel theme
// Developer Intent: Replace a workbook’s current theme with Aspose.Cells’ default theme.
// Use Cases: Standardize the look of many workbooks before publishing by applying the default theme. | Strip corporate branding from a template so end users receive a neutral workbook. | Prepare a file for external distribution to avoid theme‑related licensing issues.
// AI Prompts: Show C# code that loads an Excel file, resets its theme to the default using Aspose.Cells, and saves the output. | Create a reusable method that takes input and output paths, resets the workbook theme to Aspose.Cells’ default, and returns success status. | Explain how the CopyTheme method works in Aspose.Cells and demonstrate its use for removing custom themes.

using System;
using Aspose.Cells;

// Loads an existing .xlsx file, creates a fresh workbook that contains Aspose.Cells’ built‑in default theme, copies that theme onto the loaded workbook with CopyTheme, and saves the result. The process removes any custom theme and restores the standard default colors.
class ThemeResetExample
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Create a new workbook which contains the default theme
        Workbook defaultThemeWorkbook = new Workbook();

        // Replace the theme of the loaded workbook with the default theme
        workbook.CopyTheme(defaultThemeWorkbook);

        // Save the workbook with the refreshed theme
        workbook.Save("OutputWorkbook.xlsx");
    }
}
