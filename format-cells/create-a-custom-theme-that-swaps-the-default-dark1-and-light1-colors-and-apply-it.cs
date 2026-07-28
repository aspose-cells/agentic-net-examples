// Title: C# – Aspose.Cells: Create a Custom Theme that Swaps Dark1 (Background1) and Light1 (Text1) Colors
// Description: This example shows how to create a new Workbook, read the default Background1 and Text1 theme colors, build a 12‑element color array with those two entries swapped, apply the array as a custom theme named "SwappedDarkLightTheme", use the swapped Text1 color for a cell's font, and save the workbook as SwappedThemeDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | custom theme | swap Dark1 Light1 | Background1 Text1 | theme colors array | Workbook.CustomTheme | apply theme to cell | Excel theme programming | GitHub example | code snippet
// Common Searches: Aspose.Cells swap Dark1 and Light1 colors | create custom theme in Aspose.Cells C# | how to change workbook theme colors programmatically | apply custom theme to Excel file using Aspose.Cells | example of Workbook.CustomTheme method
// Developer Intent: Generate a custom Excel theme in Aspose.Cells that exchanges the Background1 (Dark1) and Text1 (Light1) colors and apply it to the workbook.
// Use Cases: Match corporate branding by reversing dark and light theme shades. | Demonstrate theme manipulation by styling a cell with the new Text1 color. | Distribute Excel files that carry a predefined visual style without manual formatting.
// AI Prompts: Write C# code with Aspose.Cells to create a custom theme that swaps Background1 and Text1 colors and apply it to a workbook. | Explain how to retrieve existing theme colors, construct the required 12‑element color array, and set it using Workbook.CustomTheme. | Show how to assign the swapped Text1 theme color to a cell's font after creating the custom theme.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCustomThemeDemo
{
    // This example shows how to create a new Workbook, read the default Background1 and Text1 theme colors, build a 12‑element color array with those two entries swapped, apply the array as a custom theme named "SwappedDarkLightTheme", use the swapped Text1 color for a cell's font, and save the workbook as SwappedThemeDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the current default theme colors for Background1 (Dark1) and Text1 (Light1)
            Color originalBackground1 = workbook.GetThemeColor(ThemeColorType.Background1);
            Color originalText1 = workbook.GetThemeColor(ThemeColorType.Text1);

            // Prepare a 12‑element array for the custom theme.
            // Index 0 = Background1, Index 1 = Text1, the rest keep their existing values.
            Color[] customColors = new Color[12];
            customColors[0] = originalText1;          // Swap: Background1 gets original Text1 color
            customColors[1] = originalBackground1;    // Swap: Text1 gets original Background1 color
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

            // Apply the custom theme (rule: Workbook.CustomTheme)
            workbook.CustomTheme("SwappedDarkLightTheme", customColors);

            // Demonstrate the swapped theme by applying Text1 (now dark) to a cell's font
            Cell demoCell = sheet.Cells["A1"];
            demoCell.PutValue("Swapped Theme Demo");

            Style style = workbook.CreateStyle();
            // Use the Text1 theme color for the font (which now holds the original Background1 color)
            style.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0.0);
            style.Font.Size = 14;
            demoCell.SetStyle(style);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("SwappedThemeDemo.xlsx");
        }
    }
}
