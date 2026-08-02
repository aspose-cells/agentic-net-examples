// Title: Apply a Custom Theme and Theme‑Based Styling to a New Workbook with Aspose.Cells for C#
// Description: Creates a Workbook, inserts sample data (A1:B4), defines a 12‑color custom theme named "MyCustomTheme", applies Accent1 background and Text1 font to the header, uses Accent2 background for data rows, and saves the file as CustomThemeDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells custom theme C# | apply theme colors Excel | theme‑based cell styling | C# workbook custom colors | Aspose.Cells header style | Excel theme Accent1 Accent2 | save workbook with custom theme
// Common Searches: Aspose.Cells how to create a custom theme in C# | apply theme colors to cells with Aspose.Cells .NET | style Excel header using theme Accent1 | set background color of rows with custom theme Aspose | save workbook after applying custom theme Aspose.Cells
// Developer Intent: Generate a new Excel file, define a reusable custom theme, style header and data rows with theme colors, and write the workbook to disk.
// Use Cases: Corporate report templates that automatically use brand colors defined in a custom theme. | Automated inventory or sales sheets where headers and rows are highlighted with consistent Accent colors. | Reusable Excel generators that apply a single theme to multiple workbooks, ensuring uniform styling across all outputs.
// AI Prompts: Write C# code with Aspose.Cells to create a 12‑color custom theme and apply Accent1 background to the header row. | Modify the sample to use Accent3 for data rows while keeping the Text1 font. | Explain how to extract an existing custom theme from a workbook and reuse it in another workbook with Aspose.Cells.

using Aspose.Cells;
using System.Drawing;

// Creates a Workbook, inserts sample data (A1:B4), defines a 12‑color custom theme named "MyCustomTheme", applies Accent1 background and Text1 font to the header, uses Accent2 background for data rows, and saves the file as CustomThemeDemo.xlsx using Aspose.Cells for .NET.
class CustomThemeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data
        worksheet.Cells["A1"].PutValue("Item");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["B4"].PutValue(30);

        // Define a custom theme (12 colors)
        Color[] customColors = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1
            Color.FromArgb(0, 0, 0),       // Text1
            Color.FromArgb(240, 240, 240), // Background2
            Color.FromArgb(80, 80, 80),    // Text2
            Color.FromArgb(255, 0, 0),     // Accent1
            Color.FromArgb(0, 255, 0),     // Accent2
            Color.FromArgb(0, 0, 255),     // Accent3
            Color.FromArgb(255, 165, 0),   // Accent4
            Color.FromArgb(128, 0, 128),   // Accent5
            Color.FromArgb(0, 128, 128),   // Accent6
            Color.FromArgb(0, 0, 255),     // Hyperlink
            Color.FromArgb(128, 0, 0)      // Followed Hyperlink
        };

        // Apply the custom theme
        workbook.CustomTheme("MyCustomTheme", customColors);

        // Style header row using theme colors (Accent1 background, Text1 font)
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0);
        headerStyle.Font.IsBold = true;
        headerStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
        headerStyle.Pattern = BackgroundType.Solid;
        worksheet.Cells["A1"].SetStyle(headerStyle);
        worksheet.Cells["B1"].SetStyle(headerStyle);

        // Style data rows using Accent2 background and Text1 font
        Style dataStyle = workbook.CreateStyle();
        dataStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0);
        dataStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);
        dataStyle.Pattern = BackgroundType.Solid;
        for (int row = 2; row <= 4; row++)
        {
            worksheet.Cells[$"A{row}"].SetStyle(dataStyle);
            worksheet.Cells[$"B{row}"].SetStyle(dataStyle);
        }

        // Save the workbook
        workbook.Save("CustomThemeDemo.xlsx");
    }
}
