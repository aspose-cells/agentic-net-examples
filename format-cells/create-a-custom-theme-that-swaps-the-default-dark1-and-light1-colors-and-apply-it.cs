// Title: C# – Create and Apply a Custom Theme that Swaps Background1 (Dark1) and Text1 (Light1) Colors with Aspose.Cells
// Description: This example shows how to read the default Background1 and Text1 theme colors from a workbook, exchange them in a 12‑element color array, register the array as a custom theme named "SwappedTheme", apply the new Text1 color to a cell’s font, and save the file as SwappedThemeDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells custom theme C# | swap Background1 Text1 | Dark1 Light1 theme colors | Aspose.Cells ThemeColor manipulation | Create custom theme Aspose.Cells | .NET Excel theme color swap | CustomTheme API Aspose.Cells
// Common Searches: how to swap dark1 and light1 colors in Aspose.Cells | create custom theme with swapped Background1 and Text1 in C# | apply custom theme to workbook Aspose.Cells .NET | set cell font to use swapped Text1 theme color | Aspose.Cells CustomTheme example
// Developer Intent: Create a custom Excel theme that exchanges the default Background1 (Dark1) and Text1 (Light1) colors and apply it to a workbook.
// Use Cases: Align generated Excel reports with corporate branding by swapping dark and light theme colors. | Demonstrate a custom theme by styling a cell with the newly swapped Text1 color. | Distribute workbooks that require a non‑standard color scheme without manually editing each file.
// AI Prompts: Generate C# code using Aspose.Cells to create a custom theme that swaps Background1 and Text1 colors and apply it to a workbook. | Explain how to retrieve existing theme colors, build a 12‑element color array, and register it as a custom theme in Aspose.Cells. | Show how to style a cell’s font with the swapped Text1 theme color after applying the custom theme.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCustomThemeDemo
{
    // This example shows how to read the default Background1 and Text1 theme colors from a workbook, exchange them in a 12‑element color array, register the array as a custom theme named "SwappedTheme", apply the new Text1 color to a cell’s font, and save the file as SwappedThemeDemo.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the current default theme colors for Background1 and Text1
            Color originalBackground1 = workbook.GetThemeColor(ThemeColorType.Background1);
            Color originalText1 = workbook.GetThemeColor(ThemeColorType.Text1);

            // Prepare a custom theme color array (12 entries)
            // Index 0 = Background1, Index 1 = Text1, other indices keep their existing colors
            Color[] customColors = new Color[12];
            // Swap Background1 and Text1
            customColors[0] = originalText1;      // New Background1 becomes original Text1
            customColors[1] = originalBackground1; // New Text1 becomes original Background1
            // Preserve the rest of the theme colors
            customColors[2] = workbook.GetThemeColor(ThemeColorType.Background2);
            customColors[3] = workbook.GetThemeColor(ThemeColorType.Text2);
            customColors[4] = workbook.GetThemeColor(ThemeColorType.Accent1);
            customColors[5] = workbook.GetThemeColor(ThemeColorType.Accent2);
            customColors[6] = workbook.GetThemeColor(ThemeColorType.Accent3);
            customColors[7] = workbook.GetThemeColor(ThemeColorType.Accent4);
            customColors[8] = workbook.GetThemeColor(ThemeColorType.Accent5);
            customColors[9] = workbook.GetThemeColor(ThemeColorType.Accent6);
            customColors[10] = workbook.GetThemeColor(ThemeColorType.Hyperlink);
            customColors[11] = workbook.GetThemeColor(ThemeColorType.FollowedHyperlink);

            // Apply the custom theme
            workbook.CustomTheme("SwappedTheme", customColors);

            // Demonstrate the swapped theme by applying Text1 (now dark) to a cell's font
            Cell demoCell = sheet.Cells["A1"];
            demoCell.PutValue("Swapped Theme Demo");

            Style style = workbook.CreateStyle();
            // Use the Text1 theme color for the font (which is now the original Background1 color)
            style.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0.0);
            style.Font.Size = 14;
            demoCell.SetStyle(style);

            // Save the workbook
            workbook.Save("SwappedThemeDemo.xlsx");
        }
    }
}
