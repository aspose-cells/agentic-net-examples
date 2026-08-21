// Title: Aspose.Cells .NET: Create a Custom Excel Theme and Change Accent1 & Accent2 RGB Colors
// Description: Learn how to build a 12‑color array from a workbook's default theme, replace the Accent1 and Accent2 entries with custom RGB values, register the array as a new theme using Workbook.CustomTheme, and apply the modified theme colors to cells via Style.Font.ThemeColor.
// Keywords: Aspose.Cells custom theme .NET | C# change Excel accent colors | ThemeColorType Accent1 Accent2 | Workbook.CustomTheme example | programmatic Excel theme colors | set RGB for Excel theme | apply custom theme to cells | Aspose.Cells ThemeColor usage | Excel theme customization C# | Color.FromArgb Aspose.Cells
// Common Searches: how to modify Accent1 color in Aspose.Cells | Aspose.Cells .NET custom theme example | set custom RGB values for Excel theme accents | change Excel theme colors programmatically C# | apply custom theme to workbook using Aspose.Cells
// Developer Intent: Create and apply a custom Excel theme that overrides the default Accent1 and Accent2 colors with specific RGB values using Aspose.Cells for .NET.
// Use Cases: Generate a 12‑element Color array, replace the Accent1 and Accent2 entries with Color.FromArgb, register the array as a named custom theme, and style cells to display the new accent colors. | Reuse the same custom theme across multiple worksheets or workbooks to maintain consistent branding or visual style without manually setting individual cell colors.
// AI Prompts: Show me C# code that creates a custom Excel theme in Aspose.Cells, changes Accent1 to orange and Accent2 to light blue, and applies the theme to a workbook. | Provide a step‑by‑step example of retrieving the default theme colors, updating specific accent colors, registering the custom theme, and using ThemeColor in cell styles with Aspose.Cells for .NET. | Explain how to reuse a custom theme across several worksheets in the same workbook using Aspose.Cells.

using System.Drawing;
using Aspose.Cells;

// Learn how to build a 12‑color array from a workbook's default theme, replace the Accent1 and Accent2 entries with custom RGB values, register the array as a new theme using Workbook.CustomTheme, and apply the modified theme colors to cells via Style.Font.ThemeColor.
class CustomThemeDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare an array of 12 theme colors (the required length for CustomTheme)
        Color[] themeColors = new Color[12];

        // Initialize the array with the workbook's current theme colors
        for (int i = 0; i < 12; i++)
        {
            themeColors[i] = workbook.GetThemeColor((ThemeColorType)i);
        }

        // Assign new RGB values to Accent1 (index 4) and Accent2 (index 5)
        themeColors[(int)ThemeColorType.Accent1] = Color.FromArgb(255, 128, 0);   // Orange
        themeColors[(int)ThemeColorType.Accent2] = Color.FromArgb(0, 128, 255);   // Light Blue

        // Apply the custom theme to the workbook
        workbook.CustomTheme("MyCustomTheme", themeColors);

        // Demonstrate Accent1 theme color on a cell
        Cell cellAccent1 = worksheet.Cells["A1"];
        cellAccent1.PutValue("Accent1 Text");
        Style styleAccent1 = workbook.CreateStyle();
        styleAccent1.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
        cellAccent1.SetStyle(styleAccent1);

        // Demonstrate Accent2 theme color on another cell
        Cell cellAccent2 = worksheet.Cells["A2"];
        cellAccent2.PutValue("Accent2 Text");
        Style styleAccent2 = workbook.CreateStyle();
        styleAccent2.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);
        cellAccent2.SetStyle(styleAccent2);

        // Save the workbook with the custom theme applied
        workbook.Save("CustomThemeAccentDemo.xlsx");
    }
}
