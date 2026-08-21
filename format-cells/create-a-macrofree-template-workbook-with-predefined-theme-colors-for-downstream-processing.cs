// Title: C# – Create a Macro‑Free Excel Template with a Custom Theme Using Aspose.Cells
// Description: Demonstrates how to generate a macro‑free Excel workbook, define a 12‑color custom theme, apply the theme to a cell style, and save the file as a reusable template with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# custom theme | macro free Excel workbook | programmatic theme colors | Excel template generation | apply theme to cell style | save workbook Aspose | corporate branding Excel | custom theme API | Excel styling without macros | Aspose.Cells tutorial
// Common Searches: how to create a macro‑free Excel template with Aspose.Cells | set custom theme colors in a workbook using C# | apply a custom theme to Excel cells programmatically | Aspose.Cells create reusable template workbook | define 12 theme colors in Aspose.Cells .NET
// Developer Intent: Produce a macro‑free Excel file that contains a named custom theme with predefined colors, ready for downstream automation or reporting.
// Use Cases: Generate a corporate‑branded template that downstream processes can populate without macros. | Standardize cell styling across multiple worksheets by applying a single custom theme. | Create a base workbook for automated report generation, ensuring consistent colors and formatting.
// AI Prompts: Show how to change the custom theme colors after the workbook is saved. | Provide code to copy the defined custom theme to another workbook without redefining the color array. | Explain how to combine custom theme colors with custom fonts and effects using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace TemplateWorkbookGenerator
{
    // Demonstrates how to generate a macro‑free Excel workbook, define a 12‑color custom theme, apply the theme to a cell style, and save the file as a reusable template with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro‑free by default)
            Workbook workbook = new Workbook();

            // Define 12 custom theme colors (Background1, Text1, ..., FollowedHyperlink)
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1 - White
                Color.FromArgb(0, 0, 0),       // Text1 - Black
                Color.FromArgb(242, 242, 242), // Background2 - Light Gray
                Color.FromArgb(64, 64, 64),    // Text2 - Dark Gray
                Color.FromArgb(0, 112, 192),   // Accent1 - Blue
                Color.FromArgb(255, 192, 0),   // Accent2 - Orange
                Color.FromArgb(112, 173, 71),  // Accent3 - Green
                Color.FromArgb(255, 0, 0),     // Accent4 - Red
                Color.FromArgb(255, 0, 255),   // Accent5 - Magenta
                Color.FromArgb(0, 255, 255),   // Accent6 - Cyan
                Color.FromArgb(0, 0, 255),     // Hyperlink - Blue
                Color.FromArgb(128, 0, 128)    // FollowedHyperlink - Purple
            };

            // Apply the custom theme to the workbook
            workbook.CustomTheme("MyTemplateTheme", customColors);

            // Optional: demonstrate the theme by applying Accent1 to a cell style
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Themed Cell");

            Style themedStyle = workbook.CreateStyle();
            themedStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            themedStyle.Font.Size = 12;
            cell.SetStyle(themedStyle);

            // Save the macro‑free template workbook
            workbook.Save("TemplateWorkbook.xlsx");
        }
    }
}
