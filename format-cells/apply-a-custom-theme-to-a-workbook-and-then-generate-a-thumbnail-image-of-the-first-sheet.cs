using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsThemeThumbnailDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Prepare 12 custom theme colors (Background1, Text1, Background2, Text2, Accent1‑Accent6, Hyperlink, FollowedHyperlink)
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1 - white
                Color.FromArgb(0,   0,   0),   // Text1 - black
                Color.FromArgb(240, 240, 240), // Background2 - light gray
                Color.FromArgb(80,  80,  80),  // Text2 - dark gray
                Color.FromArgb(0,   112, 192), // Accent1 - blue
                Color.FromArgb(255, 192, 0),   // Accent2 - orange
                Color.FromArgb(112, 173, 71),  // Accent3 - green
                Color.FromArgb(255, 0,   0),   // Accent4 - red
                Color.FromArgb(255, 0, 255),   // Accent5 - magenta
                Color.FromArgb(0,   255, 255), // Accent6 - cyan
                Color.FromArgb(0,   0,   255), // Hyperlink - blue
                Color.FromArgb(128, 0,   128)  // Followed Hyperlink - purple
            };

            // Apply the custom theme
            workbook.CustomTheme("MyCustomTheme", customColors);

            // Add some sample data to the first worksheet to make the thumbnail meaningful
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Custom Theme Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345);

            // Configure image rendering options (PNG thumbnail)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true // Ensure each sheet renders to a single page
            };

            // Render the first sheet (page index 0) to a PNG file
            WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);
            string thumbnailPath = "FirstSheetThumbnail.png";
            renderer.ToImage(0, thumbnailPath);

            // Save the workbook (optional, to verify the theme is persisted)
            workbook.Save("CustomThemeWorkbook.xlsx");

            Console.WriteLine($"Thumbnail generated at: {thumbnailPath}");
            Console.WriteLine("Workbook saved with custom theme.");
        }
    }
}