// Title: C# – Create a Workbook, Add a Custom Theme, and Apply Theme Colors with Aspose.Cells
// Description: Demonstrates how to instantiate a new Workbook, define a 12‑color custom theme (MyCustomTheme), assign it via Workbook.CustomTheme, fill a simple table, style the header with Accent1 background and Text1 font, style data rows with Accent2 background and Text2 font, auto‑fit columns, and save the file as CustomThemeSample.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells custom theme C# | Workbook.CustomTheme example | apply ThemeColor Aspose.Cells | theme colors styling cells | auto fit columns Aspose.Cells | C# Excel custom color palette | ThemeColorType Accent1 header | Aspose.Cells sample data styling
// Common Searches: how to create a custom theme in Aspose.Cells .NET | apply theme colors to cells using Aspose.Cells C# | Aspose.Cells example for Accent1 background and Text1 font | auto fit columns after styling with Aspose.Cells | C# code for Workbook.CustomTheme with RGB values
// Developer Intent: Create a new workbook, set a custom 12‑color theme, and style sample data using the theme’s accent and text colors.
// Use Cases: Define and apply a 12‑color custom theme to a workbook via Workbook.CustomTheme. | Style header cells with an accent background and text color from the custom theme. | Apply a different accent background and text color to data rows. | Automatically adjust column widths after applying styles. | Save the themed workbook as an XLSX file.
// AI Prompts: Generate C# code that creates a workbook, adds a custom theme with specific RGB values, and uses ThemeColor to style header and data rows in Aspose.Cells. | Explain the relationship between ThemeColor, ThemeColorType, and Workbook.CustomTheme when styling cells in Aspose.Cells. | Show how to modify the custom theme colors and refresh existing cell styles without recreating the workbook.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCustomThemeDemo
{
    // Demonstrates how to instantiate a new Workbook, define a 12‑color custom theme (MyCustomTheme), assign it via Workbook.CustomTheme, fill a simple table, style the header with Accent1 background and Text1 font, style data rows with Accent2 background and Text2 font, auto‑fit columns, and save the file as CustomThemeSample.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define 12 custom theme colors (Background1, Text1, ..., FollowedHyperlink)
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1 - White
                Color.FromArgb(0, 0, 0),       // Text1 - Black
                Color.FromArgb(240, 240, 240), // Background2 - Light Gray
                Color.FromArgb(80, 80, 80),    // Text2 - Dark Gray
                Color.FromArgb(255, 0, 0),     // Accent1 - Red
                Color.FromArgb(0, 255, 0),     // Accent2 - Green
                Color.FromArgb(0, 0, 255),     // Accent3 - Blue
                Color.FromArgb(255, 165, 0),   // Accent4 - Orange
                Color.FromArgb(128, 0, 128),   // Accent5 - Purple
                Color.FromArgb(0, 255, 255),   // Accent6 - Cyan
                Color.FromArgb(0, 0, 255),     // Hyperlink - Blue
                Color.FromArgb(128, 0, 0)      // FollowedHyperlink - Maroon
            };

            // Apply the custom theme (rule: CustomTheme)
            workbook.CustomTheme("MyCustomTheme", customColors);

            // Populate sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Cherries");
            sheet.Cells["B4"].PutValue(20);

            // Apply theme colors to header row using Accent1 for background and Text1 for font
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0);
            headerStyle.Font.IsBold = true;
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Apply theme colors to data rows (Accent2 background, Text2 font)
            Style dataStyle = workbook.CreateStyle();
            dataStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);
            dataStyle.Pattern = BackgroundType.Solid;
            dataStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Text2, 0);
            for (int row = 2; row <= 4; row++)
            {
                sheet.Cells[$"A{row}"].SetStyle(dataStyle);
                sheet.Cells[$"B{row}"].SetStyle(dataStyle);
            }

            // Auto-fit columns for better appearance
            sheet.AutoFitColumn(0);
            sheet.AutoFitColumn(1);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("CustomThemeSample.xlsx");
        }
    }
}
