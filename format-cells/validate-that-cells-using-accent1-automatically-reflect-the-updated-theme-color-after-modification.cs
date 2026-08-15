// Title: Validate that cells using Accent1 auto‑update after changing the workbook theme in Aspose.Cells (C#)
// Description: Creates a workbook, applies the Accent1 theme color to cell A1, changes the Accent1 concrete color with Workbook.SetThemeColor, and confirms the cell still references Accent1 so the new color appears when opened in Excel.
// Keywords: Aspose.Cells C# theme color | Accent1 theme update | Workbook.SetThemeColor | ForegroundThemeColor | ThemeColorType.Accent1 | Excel theme propagation | validate theme change
// Common Searches: Aspose.Cells change Accent1 theme color programmatically | Does a cell using Accent1 update after theme modification | C# verify Excel theme color propagation with Aspose.Cells | How to test theme color change effect on styled cells | Aspose.Cells theme reference validation
// Developer Intent: Ensure a cell styled with the Accent1 theme reflects the new concrete color after the workbook's Accent1 theme is modified.
// Use Cases: Apply an Accent1 style once, then alter the workbook's Accent1 color and have all styled cells update automatically. | Generate before‑and‑after Excel files to visually confirm theme color changes without re‑applying styles. | Programmatically assert that ForegroundThemeColor still points to Accent1 after a theme update for automated testing.
// AI Prompts: Write a C# unit test with Aspose.Cells that verifies a cell's ForegroundThemeColor remains Accent1 after calling Workbook.SetThemeColor. | Show code that changes the Accent1 theme color and confirms the displayed color of a cell styled with Accent1 reflects the change. | Provide a script to capture the concrete Accent1 color before and after a theme update and compare it with the cell's rendered color.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeValidation
{
    // Creates a workbook, applies the Accent1 theme color to cell A1, changes the Accent1 concrete color with Workbook.SetThemeColor, and confirms the cell still references Accent1 so the new color appears when opened in Excel.
    public class Accent1ThemeUpdateDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Step 1: Apply Accent1 theme color to a cell (A1)
                // -------------------------------------------------
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue("Accent1 Theme Test");

                // Create a style that uses the Accent1 theme color (no tint)
                Style style = workbook.CreateStyle();
                style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
                style.Pattern = BackgroundType.Solid;
                cell.SetStyle(style);

                // Save the workbook before changing the theme (optional, for inspection)
                workbook.Save("BeforeThemeChange.xlsx");

                // -------------------------------------------------
                // Step 2: Change the Accent1 theme color in the workbook
                // -------------------------------------------------
                // Set a new concrete color for Accent1 (e.g., Orange)
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.Orange);

                // -------------------------------------------------
                // Step 3: Validate that the cell reflects the updated theme color
                // -------------------------------------------------
                // Retrieve the style from the cell again
                Style updatedStyle = cell.GetStyle();

                // The ForegroundThemeColor should still reference Accent1
                ThemeColor themeRef = updatedStyle.ForegroundThemeColor;
                Console.WriteLine("Cell Foreground Theme Color Type: " + themeRef.ColorType);
                Console.WriteLine("Cell Foreground Theme Color Tint: " + themeRef.Tint);

                // Get the actual concrete color that Accent1 now maps to
                Color concreteAccent1 = workbook.GetThemeColor(ThemeColorType.Accent1);
                Console.WriteLine("Current Concrete Accent1 Color: " + concreteAccent1.Name);

                // The cell's displayed color will be the concreteAccent1 when opened in Excel.
                // Save the workbook after the theme change
                workbook.Save("AfterThemeChange.xlsx");

                // -------------------------------------------------
                // Step 4: Simple verification output
                // -------------------------------------------------
                // If the theme reference type is Accent1, we can assume the cell will reflect the updated color.
                if (themeRef.ColorType == ThemeColorType.Accent1)
                {
                    Console.WriteLine("Validation passed: Cell uses Accent1 and will reflect the updated theme color.");
                }
                else
                {
                    Console.WriteLine("Validation failed: Cell does not use Accent1 theme color.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            Accent1ThemeUpdateDemo.Run();
        }
    }
}
