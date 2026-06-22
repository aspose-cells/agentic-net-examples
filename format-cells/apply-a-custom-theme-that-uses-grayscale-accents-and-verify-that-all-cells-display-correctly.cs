using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeDemo
{
    public class GrayscaleThemeExample
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define 12 grayscale colors for the custom theme
                Color[] grayColors = new Color[]
                {
                    Color.FromArgb(255, 255, 255), // Background1 – white
                    Color.FromArgb(0, 0, 0),       // Text1 – black
                    Color.FromArgb(240, 240, 240), // Background2 – light gray
                    Color.FromArgb(64, 64, 64),    // Text2 – dark gray
                    Color.FromArgb(200, 200, 200), // Accent1 – medium light gray
                    Color.FromArgb(180, 180, 180), // Accent2 – medium gray
                    Color.FromArgb(160, 160, 160), // Accent3 – medium dark gray
                    Color.FromArgb(140, 140, 140), // Accent4 – dark gray
                    Color.FromArgb(120, 120, 120), // Accent5 – darker gray
                    Color.FromArgb(100, 100, 100), // Accent6 – very dark gray
                    Color.FromArgb(30, 144, 255),  // Hyperlink – default blue
                    Color.FromArgb(128, 0, 128)    // Followed Hyperlink – default purple
                };

                // Apply the custom grayscale theme
                workbook.CustomTheme("GrayTheme", grayColors);

                // Cell A1 – Font using Accent1
                Cell cellA1 = sheet.Cells["A1"];
                cellA1.PutValue("Accent1 Font");
                Style styleA1 = workbook.CreateStyle();
                styleA1.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
                cellA1.SetStyle(styleA1);

                // Cell A2 – Background using Accent2
                Cell cellA2 = sheet.Cells["A2"];
                cellA2.PutValue("Accent2 Background");
                Style styleA2 = workbook.CreateStyle();
                styleA2.Pattern = BackgroundType.Solid;
                styleA2.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);
                cellA2.SetStyle(styleA2);

                // Cell A3 – Font using Accent3 with 30% lighten tint
                Cell cellA3 = sheet.Cells["A3"];
                cellA3.PutValue("Accent3 Font + Tint");
                Style styleA3 = workbook.CreateStyle();
                styleA3.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0.3);
                cellA3.SetStyle(styleA3);

                // Cell A4 – Background using Accent4 with 20% darken tint
                Cell cellA4 = sheet.Cells["A4"];
                cellA4.PutValue("Accent4 Background + Tint");
                Style styleA4 = workbook.CreateStyle();
                styleA4.Pattern = BackgroundType.Solid;
                styleA4.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent4, -0.2);
                cellA4.SetStyle(styleA4);

                // Verification output
                Console.WriteLine("Verification of applied theme colors:");
                Console.WriteLine($"A1 Font Theme: {cellA1.GetStyle().Font.ThemeColor.ColorType}, Tint={cellA1.GetStyle().Font.ThemeColor.Tint}");
                Console.WriteLine($"A2 Background Theme: {cellA2.GetStyle().ForegroundThemeColor.ColorType}, Tint={cellA2.GetStyle().ForegroundThemeColor.Tint}");
                Console.WriteLine($"A3 Font Theme: {cellA3.GetStyle().Font.ThemeColor.ColorType}, Tint={cellA3.GetStyle().Font.ThemeColor.Tint}");
                Console.WriteLine($"A4 Background Theme: {cellA4.GetStyle().ForegroundThemeColor.ColorType}, Tint={cellA4.GetStyle().ForegroundThemeColor.Tint}");

                // Save the workbook
                string outputPath = "GrayscaleThemeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            GrayscaleThemeExample.Run();
        }
    }
}