using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemePdfPreview
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook from file
            Workbook workbook = new Workbook("input.xlsx");

            // Define 12 custom theme colors (Background1, Text1, Background2, Text2, Accent1‑Accent6, Hyperlink, FollowedHyperlink)
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1 - White
                Color.FromArgb(0, 0, 0),       // Text1 - Black
                Color.FromArgb(240, 240, 240), // Background2 - Light gray
                Color.FromArgb(80, 80, 80),    // Text2 - Dark gray
                Color.FromArgb(0, 120, 215),   // Accent1 - Blue
                Color.FromArgb(0, 153, 0),     // Accent2 - Green
                Color.FromArgb(255, 185, 0),   // Accent3 - Orange
                Color.FromArgb(255, 0, 0),     // Accent4 - Red
                Color.FromArgb(112, 48, 160),  // Accent5 - Purple
                Color.FromArgb(255, 192, 0),   // Accent6 - Gold
                Color.FromArgb(0, 0, 255),     // Hyperlink - Blue
                Color.FromArgb(128, 0, 128)    // FollowedHyperlink - Purple
            };

            // Apply the custom theme to the workbook
            workbook.CustomTheme("MyCustomTheme", customColors);

            // OPTIONAL: Demonstrate the theme by applying an accent color to a cell style
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Custom Theme Applied");

            Style style = workbook.CreateStyle();
            style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            style.Font.Size = 14;
            cell.SetStyle(style);

            // Save the first worksheet as a PDF file (acts as a preview)
            // By default, Save will include all worksheets; to keep only the first sheet,
            // we can hide the others before saving.
            for (int i = 1; i < workbook.Worksheets.Count; i++)
            {
                workbook.Worksheets[i].IsVisible = false;
            }

            workbook.Save("FirstSheetPreview.pdf", SaveFormat.Pdf);

            Console.WriteLine("PDF preview of the first sheet has been generated.");
        }
    }
}