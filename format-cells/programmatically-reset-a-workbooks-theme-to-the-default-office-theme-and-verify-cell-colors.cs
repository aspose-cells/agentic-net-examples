// Title: Reset workbook theme to default Office theme and verify colors with Aspose.Cells (C#)
// Description: Demonstrates how to apply a custom theme, then programmatically revert a workbook to the built‑in Office theme using Aspose.Cells' CopyTheme method, verify the reset with GetThemeColor, and save the result.
// Keywords: Aspose.Cells Reset Theme | CopyTheme method | GetThemeColor | default Office theme | C# Aspose.Cells example | theme color verification | Aspose.Cells workbook theme reset
// Common Searches: How to reset a workbook theme to default using Aspose.Cells C# | Aspose.Cells CopyTheme example | Verify theme colors after reset Aspose.Cells | Reset custom theme to Office default Aspose.Cells .NET | GetThemeColor usage Aspose.Cells
// Developer Intent: Programmatically revert a workbook’s custom theme to the built‑in Office theme and confirm that the theme colors match the default.
// Use Cases: Revert a styled workbook to the standard Office theme before distribution | Automated testing to ensure no custom theme colors remain after processing | Create a template workbook with the default theme by copying from a fresh workbook | Batch process multiple files to strip custom themes and retain default styling
// AI Prompts: Provide C# code using Aspose.Cells to copy the default Office theme into an existing workbook and validate the Accent1 color with GetThemeColor. | Show how to reset a custom theme to the built‑in Office theme in Aspose.Cells and compare all theme colors. | Explain step‑by‑step how CopyTheme and GetThemeColor work together to verify a theme reset in .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeResetDemo
{
    // Demonstrates how to apply a custom theme, then programmatically revert a workbook to the built‑in Office theme using Aspose.Cells' CopyTheme method, verify the reset with GetThemeColor, and save the result.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a workbook and apply a custom theme
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet ws = workbook.Worksheets[0];

            // Define 12 custom theme colors (example values)
            Color[] customColors = new Color[]
            {
                Color.Red,          // Background1
                Color.Green,        // Text1
                Color.Blue,         // Background2
                Color.Yellow,       // Text2
                Color.Magenta,      // Accent1
                Color.Cyan,         // Accent2
                Color.Purple,       // Accent3
                Color.Olive,        // Accent4
                Color.Teal,         // Accent5
                Color.Maroon,       // Accent6
                Color.DarkGreen,    // Hyperlink
                Color.Navy          // FollowedHyperlink
            };

            // Apply the custom theme
            workbook.CustomTheme("MyCustomTheme", customColors);

            // Create a cell that uses the Accent1 theme color
            Cell themedCell = ws.Cells["A1"];
            themedCell.PutValue("Custom Theme Cell");
            Style themedStyle = workbook.CreateStyle();
            themedStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            themedCell.SetStyle(themedStyle);

            // Save the workbook with the custom theme (optional)
            workbook.Save("CustomThemeWorkbook.xlsx");

            // -------------------------------------------------
            // 2. Reset the workbook's theme to the default Office theme
            // -------------------------------------------------
            // Create a fresh workbook that contains the default Office theme
            Workbook defaultThemeWorkbook = new Workbook(); // default theme

            // Copy the default theme into the original workbook
            workbook.CopyTheme(defaultThemeWorkbook);

            // -------------------------------------------------
            // 3. Verify that the theme has been reset
            // -------------------------------------------------
            // Retrieve the Accent1 color from both workbooks
            Color accent1AfterReset = workbook.GetThemeColor(ThemeColorType.Accent1);
            Color accent1Default = defaultThemeWorkbook.GetThemeColor(ThemeColorType.Accent1);

            // Output verification result
            Console.WriteLine($"Accent1 after reset:   A={accent1AfterReset.A}, R={accent1AfterReset.R}, G={accent1AfterReset.G}, B={accent1AfterReset.B}");
            Console.WriteLine($"Accent1 default theme: A={accent1Default.A}, R={accent1Default.R}, G={accent1Default.G}, B={accent1Default.B}");
            Console.WriteLine($"Theme reset successful: {accent1AfterReset.Equals(accent1Default)}");

            // -------------------------------------------------
            // 4. Save the workbook after resetting the theme
            // -------------------------------------------------
            workbook.Save("ResetToDefaultTheme.xlsx");
        }
    }
}
