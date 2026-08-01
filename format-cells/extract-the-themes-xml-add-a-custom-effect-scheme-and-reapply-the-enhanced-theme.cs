// Title: Create and Apply a Custom Effect Scheme with Aspose.Cells for .NET (C#)
// Description: This example shows how to read the current workbook theme name, define a 12‑color custom effect scheme, replace the existing theme using Aspose.Cells' CustomTheme method, style a cell with an Accent theme color, and save the workbook with the new theme applied.
// Keywords: Aspose.Cells | C# | .NET | CustomTheme | workbook theme | effect scheme | ThemeColor | Accent1 | Excel theme customization | color palette | cell styling
// Common Searches: Aspose.Cells apply custom theme colors C# | How to change workbook theme with Aspose.Cells | Create custom effect scheme in Aspose.Cells .NET | Set cell font using ThemeColor Accent1 after custom theme | Retrieve current theme name Aspose.Cells workbook
// Developer Intent: Define a custom 12‑color effect scheme and apply it as a new theme to an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Replace the default Excel theme with a brand‑specific color palette via CustomTheme. | Demonstrate the new theme by formatting a cell’s font with an Accent theme color. | Generate workbooks that retain the custom theme when opened in Microsoft Excel.
// AI Prompts: Write C# code that extracts the current theme name from an Aspose.Cells workbook and applies a custom effect scheme using a 12‑color array. | Show how to set a cell’s font color to ThemeColor Accent3 after a custom theme has been applied with Aspose.Cells. | Explain how to access the internal theme XML of an Excel file using Aspose.Cells package parts when the public API does not expose it.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeEnhancement
{
    // This example shows how to read the current workbook theme name, define a 12‑color custom effect scheme, replace the existing theme using Aspose.Cells' CustomTheme method, style a cell with an Accent theme color, and save the workbook with the new theme applied.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // lifecycle: create

            // Access the first worksheet for demonstration purposes
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // Step 1: Extract current theme information
            // -----------------------------------------------------------------
            // Aspose.Cells provides the theme name via the Theme property.
            // Direct access to the theme's XML is not exposed through the API,
            // so we capture the theme name as the available metadata.
            string currentThemeName = workbook.Theme;
            Console.WriteLine("Current theme name: " + currentThemeName);

            // Placeholder: If XML extraction were required, it would involve
            // accessing the internal theme part of the workbook package.
            // This is not available via the public API, so we proceed with
            // enhancing the theme using the supported methods.

            // -----------------------------------------------------------------
            // Step 2: Define a custom effect scheme (custom colors)
            // -----------------------------------------------------------------
            // The CustomTheme method expects an array of 12 colors that map to
            // the standard theme slots (Background1, Text1, ..., FollowedHyperlink).
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1 (white)
                Color.FromArgb(0, 0, 0),       // Text1 (black)
                Color.FromArgb(240, 240, 240), // Background2 (light gray)
                Color.FromArgb(80, 80, 80),    // Text2 (dark gray)
                Color.FromArgb(255, 99, 71),   // Accent1 (tomato)
                Color.FromArgb(60, 179, 113),  // Accent2 (medium sea green)
                Color.FromArgb(30, 144, 255),  // Accent3 (dodger blue)
                Color.FromArgb(255, 215, 0),   // Accent4 (gold)
                Color.FromArgb(218, 112, 214), // Accent5 (orchid)
                Color.FromArgb(255, 165, 0),   // Accent6 (orange)
                Color.FromArgb(0, 0, 255),     // Hyperlink (blue)
                Color.FromArgb(255, 0, 0)      // FollowedHyperlink (red)
            };

            // -----------------------------------------------------------------
            // Step 3: Apply the custom theme to the workbook
            // -----------------------------------------------------------------
            // The CustomTheme method replaces the existing theme with the
            // provided color scheme while preserving the theme name.
            workbook.CustomTheme("EnhancedCustomTheme", customColors);

            // Verify that the theme name has been updated
            Console.WriteLine("New theme name after customization: " + workbook.Theme);

            // -----------------------------------------------------------------
            // Step 4: Demonstrate the applied theme on a cell
            // -----------------------------------------------------------------
            Cell demoCell = sheet.Cells["A1"];
            demoCell.PutValue("Theme with Custom Effect Scheme");
            Style demoStyle = workbook.CreateStyle();
            // Use Accent1 from the custom theme for the font color
            demoStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            demoStyle.Font.Size = 14;
            demoCell.SetStyle(demoStyle);

            // -----------------------------------------------------------------
            // Step 5: Save the enhanced workbook
            // -----------------------------------------------------------------
            workbook.Save("EnhancedThemeWorkbook.xlsx"); // lifecycle: save
        }
    }
}
