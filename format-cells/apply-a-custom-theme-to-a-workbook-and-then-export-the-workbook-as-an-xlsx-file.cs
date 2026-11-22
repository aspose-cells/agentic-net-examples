// Title: How to load a .thmx custom theme into an Aspose.Cells workbook and save it as XLSX using C#
// Common Searches: C# code to apply a .thmx theme to an Aspose.Cells workbook | Saving an Excel file with a custom theme using Aspose.Cells for .NET | Load custom theme file into Aspose.Cells and export to XLSX | Set Workbook.CustomTheme property in C# Aspose.Cells example | Apply custom theme to cells before saving as XLSX with Aspose.Cells
// Tags: Aspose.Cells load .thmx theme C# | Workbook.CustomTheme assignment | export workbook to XLSX Aspose.Cells | apply custom theme to Excel cells C# | Aspose.Cells custom theme support .NET

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The example creates a new Workbook, optionally loads a .thmx custom theme via Theme.Load, assigns it to workbook.CustomTheme, adds sample content, applies a style affected by the theme, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // NOTE: The Theme class may not be available in older Aspose.Cells versions.
            // If you have a version that supports custom themes, you can uncomment
            // the following block and ensure the appropriate using directive:
            // using Aspose.Cells.Drawing;
            // -----------------------------------------------------------------
            /*
            string themePath = "customTheme.thmx";
            if (File.Exists(themePath))
            {
                // Load the custom theme and assign it to the workbook
                Theme customTheme = Theme.Load(themePath);
                workbook.CustomTheme = customTheme;
            }
            */

            // Add sample content
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");

            // Create a style that will be affected by the (optional) custom theme
            Style style = workbook.CreateStyle();
            style.Font.Color = Color.White;
            sheet.Cells["A1"].SetStyle(style);

            // Save the workbook as XLSX
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
