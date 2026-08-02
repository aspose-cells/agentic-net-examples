// Title: Clone Excel Theme from a Template Workbook with Aspose.Cells CopyTheme (C#)
// Description: C# code that loads a template Excel file containing a predefined theme, creates a new workbook, clones the theme using Aspose.Cells' CopyTheme method, and saves the themed workbook. Ideal for maintaining consistent branding across generated spreadsheets.
// Keywords: Aspose.Cells | CopyTheme | C# | Excel theme clone | template workbook theme | copy workbook theme .NET | Aspose.Cells example | theme transfer Excel
// Common Searches: Aspose.Cells copy theme C# | CopyTheme method example | clone Excel theme programmatically | apply template theme to new workbook Aspose | transfer Excel theme between workbooks .NET
// Developer Intent: Copy the theme from an existing template workbook and assign it to a newly created workbook.
// Use Cases: Generate a series of reports that share corporate branding by cloning the theme from a master template. | Automate creation of department‑specific spreadsheets, ensuring each new file inherits the predefined style set. | Run a batch process that produces client‑customized workbooks, all using a consistent theme copied from a source file.
// AI Prompts: Provide a C# snippet that uses Aspose.Cells to copy a theme from a template workbook to a new workbook and saves the output. | Explain which theme elements (fonts, colors, effects) are transferred when using the CopyTheme method in Aspose.Cells. | Show how to add error handling for missing template files while cloning a theme with Aspose.Cells.

using System;
using Aspose.Cells;

namespace ThemeCloneExample
{
    // C# code that loads a template Excel file containing a predefined theme, creates a new workbook, clones the theme using Aspose.Cells' CopyTheme method, and saves the themed workbook. Ideal for maintaining consistent branding across generated spreadsheets.
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains the desired theme
            // Replace "template.xlsx" with the actual path to your template file
            Workbook templateWorkbook = new Workbook("template.xlsx");

            // Create a new empty workbook
            Workbook newWorkbook = new Workbook();

            // Clone the theme from the template workbook to the new workbook
            newWorkbook.CopyTheme(templateWorkbook);

            // Save the new workbook with the cloned theme
            // Replace "cloned_theme_output.xlsx" with your desired output file name/path
            newWorkbook.Save("cloned_theme_output.xlsx");

            Console.WriteLine("Theme cloned successfully and workbook saved.");
        }
    }
}
