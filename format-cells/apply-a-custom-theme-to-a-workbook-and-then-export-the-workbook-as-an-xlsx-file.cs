// Title: Aspose.Cells C# – Apply a Custom Theme to a Workbook and Export as XLSX
// Description: This example demonstrates how to create a new Workbook with Aspose.Cells for .NET, define a 12‑color custom theme, apply the theme to the workbook, style a cell using a theme accent, and save the result as an XLSX file.
// Keywords: Aspose.Cells C# custom theme | Workbook.CustomTheme API | apply Excel theme colors Aspose | save workbook as XLSX .NET | Excel theme accent1 Aspose.Cells | C# Excel styling example | Aspose.Cells for .NET code sample | custom Excel palette Aspose
// Common Searches: how to set a custom theme in Aspose.Cells C# | Aspose.Cells apply theme colors and save as xlsx | C# example for Workbook.CustomTheme method | export Excel file with custom palette using Aspose | Aspose.Cells theme accent color usage
// Developer Intent: Create a workbook, define and apply a custom color theme, style cells with theme accents, and generate an XLSX file.
// Use Cases: Brand automated reports with a corporate color palette by applying a custom theme before export. | Build a reusable Excel template where all formatting references theme colors for consistent styling across multiple outputs. | Showcase the visual impact of theme colors in training material or documentation using a demo workbook.
// AI Prompts: Generate C# code that reads theme colors from a JSON file and applies them to an Aspose.Cells workbook. | Explain how to propagate a custom theme to every worksheet in a multi‑sheet workbook and keep the theme after saving to XLSX. | Provide a snippet to extract and list the custom theme colors from an existing XLSX file using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCustomThemeDemo
{
    // This example demonstrates how to create a new Workbook with Aspose.Cells for .NET, define a 12‑color custom theme, apply the theme to the workbook, style a cell using a theme accent, and save the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Initialize a new workbook (uses the Workbook() constructor)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a sample cell to show the effect of the custom theme
            Cell demoCell = sheet.Cells["A1"];
            demoCell.PutValue("Custom Theme Demo");
            Style demoStyle = workbook.CreateStyle();
            demoStyle.Font.IsBold = true;
            demoStyle.Font.Size = 14;
            // Use a theme color (Accent1) which will reflect the custom theme later
            demoStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
            demoCell.SetStyle(demoStyle);

            // Define 12 custom theme colors (Background1, Text1, ..., Followed Hyperlink)
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1 - White
                Color.FromArgb(0, 0, 0),       // Text1 - Black
                Color.FromArgb(240, 240, 240), // Background2 - Light Gray
                Color.FromArgb(80, 80, 80),    // Text2 - Dark Gray
                Color.FromArgb(0, 112, 192),   // Accent1 - Blue
                Color.FromArgb(255, 192, 0),   // Accent2 - Orange
                Color.FromArgb(112, 173, 71),  // Accent3 - Green
                Color.FromArgb(255, 0, 0),     // Accent4 - Red
                Color.FromArgb(255, 0, 255),   // Accent5 - Magenta
                Color.FromArgb(0, 255, 255),   // Accent6 - Cyan
                Color.FromArgb(0, 0, 255),     // Hyperlink - Blue
                Color.FromArgb(128, 0, 128)    // Followed Hyperlink - Purple
            };

            // Apply the custom theme to the workbook (uses Workbook.CustomTheme)
            workbook.CustomTheme("MyCustomTheme", customColors);

            // Save the workbook as XLSX (uses Workbook.Save with SaveFormat)
            workbook.Save("CustomThemeWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
