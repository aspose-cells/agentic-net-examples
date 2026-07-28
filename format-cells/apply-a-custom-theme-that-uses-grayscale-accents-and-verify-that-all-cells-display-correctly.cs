// Title: Apply a Custom Grayscale Theme in Aspose.Cells (C#) and Validate Cell Styling
// Description: Demonstrates how to create a 12‑color grayscale palette, register it with Workbook.CustomTheme, and use ThemeColor accents for font, fill, and border styling in an Aspose.Cells workbook before saving as GrayThemeDemo.xlsx.
// Keywords: Aspose.Cells custom theme C# | grayscale Excel theme programmatically | Workbook.CustomTheme example | ThemeColor Accent1 Accent2 Accent3 | apply theme colors to cells Aspose | verify Excel theme styling | C# Aspose.Cells styling guide
// Common Searches: how to create a custom grayscale theme with Aspose.Cells | apply theme accents to font fill border in C# Excel | Aspose.Cells CustomTheme usage example | verify theme colors in generated workbook | C# code for grayscale Excel theme Aspose
// Developer Intent: Create a grayscale custom theme, apply its accent colors to cell font, background, and border, and save the workbook to confirm visual results.
// Use Cases: Define a 12‑color grayscale array and register it via workbook.CustomTheme("GrayTheme", colors). | Set a cell’s Font.ThemeColor to Accent1 to display grayscale text. | Apply Accent2 as the cell’s background fill using ForegroundThemeColor. | Style a cell’s bottom border with Accent3 via Borders[BorderType.BottomBorder].ThemeColor. | Save the workbook and open the .xlsx file to ensure the theme renders correctly.
// AI Prompts: Generate C# code that builds a 12‑color grayscale theme and registers it with Aspose.Cells. | Show how to assign ThemeColor Accent1 to a cell’s font, Accent2 to its fill, and Accent3 to its border using Aspose.Cells styles. | Provide a method to programmatically compare a cell’s actual ThemeColor values with the defined grayscale palette after saving.

using Aspose.Cells;
using System.Drawing;

// Demonstrates how to create a 12‑color grayscale palette, register it with Workbook.CustomTheme, and use ThemeColor accents for font, fill, and border styling in an Aspose.Cells workbook before saving as GrayThemeDemo.xlsx.
class CustomThemeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data to demonstrate the theme
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Data 1");
        sheet.Cells["A3"].PutValue("Data 2");

        // Define a grayscale theme (12 colors as required)
        Color[] grayTheme = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1 – white
            Color.FromArgb(0,   0,   0),   // Text1 – black
            Color.FromArgb(242, 242, 242), // Background2 – light gray
            Color.FromArgb(64,  64,  64),  // Text2 – dark gray
            Color.FromArgb(200, 200, 200), // Accent1
            Color.FromArgb(180, 180, 180), // Accent2
            Color.FromArgb(160, 160, 160), // Accent3
            Color.FromArgb(140, 140, 140), // Accent4
            Color.FromArgb(120, 120, 120), // Accent5
            Color.FromArgb(100, 100, 100), // Accent6
            Color.FromArgb(0,   0, 255),   // Hyperlink – keep blue for visibility
            Color.FromArgb(128, 0, 128)    // Followed Hyperlink – purple
        };

        // Apply the custom grayscale theme
        workbook.CustomTheme("GrayTheme", grayTheme);

        // Verify theme colors by applying them to cells

        // Cell B1 – font uses Accent1
        Style style1 = workbook.CreateStyle();
        style1.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
        sheet.Cells["B1"].PutValue("Accent1 Text");
        sheet.Cells["B1"].SetStyle(style1);

        // Cell B2 – background uses Accent2
        Style style2 = workbook.CreateStyle();
        style2.Pattern = BackgroundType.Solid;
        style2.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);
        sheet.Cells["B2"].PutValue("Accent2 Fill");
        sheet.Cells["B2"].SetStyle(style2);

        // Cell B3 – bottom border uses Accent3
        Style style3 = workbook.CreateStyle();
        style3.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        style3.Borders[BorderType.BottomBorder].ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);
        sheet.Cells["B3"].PutValue("Accent3 Border");
        sheet.Cells["B3"].SetStyle(style3);

        // Save the workbook to verify the visual result
        workbook.Save("GrayThemeDemo.xlsx");
    }
}
