// Title: Clone Excel Theme from a Template Workbook with Aspose.Cells for .NET
// Description: Loads a template file that contains the desired theme, creates a blank workbook, copies the theme using the CopyTheme method, and saves the new workbook with the transferred styling.
// Keywords: Aspose.Cells CopyTheme | copy Excel theme .NET | apply workbook theme programmatically | transfer Excel theme between workbooks | clone Excel theme Aspose
// Common Searches: Aspose.Cells copy theme example | how to transfer Excel theme using .NET | copy theme from one workbook to another Aspose | clone workbook theme programmatically
// Developer Intent: Assign the theme from an existing template workbook to a newly created workbook.
// Use Cases: Generate blank reports that automatically inherit corporate branding by reusing a master template's theme. | Produce multiple departmental spreadsheets with uniform styling without manually formatting each file. | Automate creation of client‑specific workbooks that must match a supplied template's visual theme.
// AI Prompts: Show how to copy a theme from a template workbook to a new workbook using Aspose.Cells for .NET, including error handling for missing files. | Provide a code sample that clones a theme and then adds custom cell formatting after the theme is applied. | Explain the differences between using CopyTheme and manually constructing a Theme object in Aspose.Cells.

using System;
using Aspose.Cells;

namespace ThemeCloneDemo
{
    // Loads a template file that contains the desired theme, creates a blank workbook, copies the theme using the CopyTheme method, and saves the new workbook with the transferred styling.
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains the desired theme
            Workbook templateWorkbook = new Workbook("Template.xlsx");

            // Create a new empty workbook
            Workbook newWorkbook = new Workbook();

            // Clone the theme from the template workbook to the new workbook
            newWorkbook.CopyTheme(templateWorkbook);

            // Save the new workbook with the cloned theme
            newWorkbook.Save("ClonedThemeWorkbook.xlsx");
        }
    }
}
