using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class CustomThemeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Custom Theme Demo");
            sheet.Cells["A2"].PutValue("Header");
            sheet.Cells["A3"].PutValue("Data 1");
            sheet.Cells["B3"].PutValue(123);
            sheet.Cells["A4"].PutValue("Data 2");
            sheet.Cells["B4"].PutValue(456);

            // Define 12 custom theme colors
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1
                Color.FromArgb(0, 0, 0),       // Text1
                Color.FromArgb(240, 240, 240), // Background2
                Color.FromArgb(80, 80, 80),    // Text2
                Color.FromArgb(0, 120, 215),   // Accent1
                Color.FromArgb(0, 153, 0),     // Accent2
                Color.FromArgb(255, 185, 0),   // Accent3
                Color.FromArgb(255, 0, 0),     // Accent4
                Color.FromArgb(112, 48, 160),  // Accent5
                Color.FromArgb(255, 192, 0),   // Accent6
                Color.FromArgb(0, 0, 255),     // Hyperlink
                Color.FromArgb(128, 0, 128)    // FollowedHyperlink
            };

            // Apply the custom theme
            workbook.CustomTheme("MyCustomTheme", customColors);

            // Style the header using Accent1 for font and Background2 for fill
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Size = 14;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Background2, 0);
            sheet.Cells["A2"].SetStyle(headerStyle);

            // Style data rows: Accent2 for font, Accent3 for background
            Style dataStyle = workbook.CreateStyle();
            dataStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);
            dataStyle.Font.Size = 12;
            dataStyle.Pattern = BackgroundType.Solid;
            dataStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);

            // Apply style to the data range A3:B4
            Aspose.Cells.Range dataRange = sheet.Cells.CreateRange("A3:B4");
            dataRange.ApplyStyle(dataStyle, new StyleFlag { Font = true, CellShading = true });

            // Save the workbook (ensure the directory exists)
            string outputPath = "CustomThemeDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}