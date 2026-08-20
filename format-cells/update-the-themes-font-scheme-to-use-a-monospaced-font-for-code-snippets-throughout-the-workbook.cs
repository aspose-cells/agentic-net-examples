// Title: Set Workbook Theme to a Monospaced Font (Consolas) Using Aspose.Cells for .NET
// Description: Demonstrates how to modify the DefaultStyle of an Aspose.Cells Workbook to apply a monospaced font across the entire workbook, ensuring all cells inherit the Consolas typeface—ideal for displaying code snippets. The example creates a workbook, updates the font name and size, reassigns the style, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | Excel workbook theme | default style font | monospaced font | Consolas | code snippets in Excel | set workbook font | change default font Aspose
// Common Searches: how to change default font in Aspose.Cells workbook | apply monospaced font to all cells with Aspose.Cells .NET | set Consolas as workbook theme font using C# | Aspose.Cells default style font size and name | global font change for generated Excel file
// Developer Intent: Apply a monospaced font to the workbook’s theme so every cell inherits the same fixed‑width typeface.
// Use Cases: Generating Excel reports that embed source code blocks. | Standardizing font appearance for documentation workbooks. | Ensuring consistent visual style for data exported from applications.
// AI Prompts: Show C# code that sets the workbook DefaultStyle font to Consolas with Aspose.Cells. | Explain how to change the theme font for an entire Excel file using Aspose.Cells for .NET. | Provide a step‑by‑step guide to apply a monospaced font to all cells in a newly created workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsThemeFontUpdate
{
    // Demonstrates how to modify the DefaultStyle of an Aspose.Cells Workbook to apply a monospaced font across the entire workbook, ensuring all cells inherit the Consolas typeface—ideal for displaying code snippets. The example creates a workbook, updates the font name and size, reassigns the style, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Update the default style of the workbook to use a monospaced font.
            // This will affect all cells that do not have an explicit style applied,
            // effectively applying the monospaced font throughout the workbook.
            Style defaultStyle = workbook.DefaultStyle;
            defaultStyle.Font.Name = "Consolas";   // Monospaced font
            defaultStyle.Font.Size = 11;           // Typical size for code snippets
            // Optionally, you can set the scheme type if needed
            // defaultStyle.Font.SchemeType = FontSchemeType.Minor;

            // Assign the modified style back to the workbook
            workbook.DefaultStyle = defaultStyle;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("WorkbookWithMonospacedTheme.xlsx");
        }
    }
}
