// Title: Aspose.Cells C# – Change Workbook Accent4 Theme Color from User Input and Refresh Dependent Styles
// Description: A complete C# example that creates a workbook, adds cells using the Accent4 theme for font and background, reads a hex color from the console, updates the workbook’s Accent4 theme via Workbook.SetThemeColor, iterates every worksheet to reapply styles that reference Accent4 (preserving tints), and saves the file as UpdatedAccent4Theme.xlsx.
// Keywords: Aspose.Cells C# theme color | set Accent4 color programmatically | refresh theme‑dependent styles | Workbook.SetThemeColor example | user‑defined Excel theme color .NET | update Excel accent color | Aspose.Cells GitHub sample
// Common Searches: how to change Accent4 theme color Aspose.Cells | refresh cells after SetThemeColor | C# ask user for hex color Excel theme | apply custom theme accent in .NET | Aspose.Cells update workbook theme color
// Developer Intent: Modify the workbook’s Accent4 theme color based on a hex value supplied at runtime and ensure every cell that uses that accent reflects the new color.
// Use Cases: Brand‑aware templates where the corporate accent color can be chosen by the end‑user. | Reporting engines that generate Excel files with client‑specific theme colors. | Batch processing of multiple workbooks to apply a new corporate Accent4 shade and automatically refresh all dependent formatting.
// AI Prompts: Generate C# code using Aspose.Cells that prompts for a hex color, sets the Accent4 theme, and updates all styles referencing it. | Explain why re‑applying cell styles is necessary after calling Workbook.SetThemeColor for an accent color. | Provide a step‑by‑step guide to validate hex input, convert it to System.Drawing.Color, and apply it to an existing Excel workbook with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace UpdateAccent4Theme
{
    // A complete C# example that creates a workbook, adds cells using the Accent4 theme for font and background, reads a hex color from the console, updates the workbook’s Accent4 theme via Workbook.SetThemeColor, iterates every worksheet to reapply styles that reference Accent4 (preserving tints), and saves the file as UpdatedAccent4Theme.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Step 1: Add sample cells that use the Accent4 theme color
            // ------------------------------------------------------------
            // Cell A1 – font uses Accent4
            Cell cellA1 = cells["A1"];
            cellA1.PutValue("Accent4 Font");
            Style styleA1 = cellA1.GetStyle();
            styleA1.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent4, 0); // no tint
            cellA1.SetStyle(styleA1);

            // Cell A2 – background uses Accent4
            Cell cellA2 = cells["A2"];
            cellA2.PutValue("Accent4 Background");
            Style styleA2 = cellA2.GetStyle();
            styleA2.Pattern = BackgroundType.Solid;
            styleA2.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent4, 0);
            cellA2.SetStyle(styleA2);

            // ------------------------------------------------------------
            // Step 2: Ask user for new Accent4 color (RGB format)
            // ------------------------------------------------------------
            Console.WriteLine("Enter new Accent4 color in hex format (e.g., FF3366):");
            string hex = Console.ReadLine()?.Trim();

            // Validate and convert hex to Color
            if (string.IsNullOrEmpty(hex) || hex.Length != 6)
            {
                Console.WriteLine("Invalid input. Using default color (Blue).");
                hex = "0000FF";
            }

            int r = Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = Convert.ToInt32(hex.Substring(4, 2), 16);
            Color newAccent4 = Color.FromArgb(r, g, b);

            // ------------------------------------------------------------
            // Step 3: Update the workbook theme's Accent4 color
            // ------------------------------------------------------------
            workbook.SetThemeColor(ThemeColorType.Accent4, newAccent4);

            // ------------------------------------------------------------
            // Step 4: Refresh all cell styles that depend on Accent4
            // ------------------------------------------------------------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Cell cell in ws.Cells)
                {
                    // Retrieve current style
                    Style curStyle = cell.GetStyle();

                    bool needsRefresh = false;

                    // Check font theme color
                    if (curStyle.Font.ThemeColor != null &&
                        curStyle.Font.ThemeColor.ColorType == ThemeColorType.Accent4)
                    {
                        // Reassign to trigger refresh (tint unchanged)
                        curStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent4, curStyle.Font.ThemeColor.Tint);
                        needsRefresh = true;
                    }

                    // Check foreground theme color
                    if (curStyle.ForegroundThemeColor != null &&
                        curStyle.ForegroundThemeColor.ColorType == ThemeColorType.Accent4)
                    {
                        curStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent4, curStyle.ForegroundThemeColor.Tint);
                        needsRefresh = true;
                    }

                    // Check background theme color
                    if (curStyle.BackgroundThemeColor != null &&
                        curStyle.BackgroundThemeColor.ColorType == ThemeColorType.Accent4)
                    {
                        curStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent4, curStyle.BackgroundThemeColor.Tint);
                        needsRefresh = true;
                    }

                    // If any theme‑dependent property was found, reapply the style
                    if (needsRefresh)
                    {
                        cell.SetStyle(curStyle);
                    }
                }
            }

            // ------------------------------------------------------------
            // Step 5: Save the workbook (lifecycle rule: save)
            // ------------------------------------------------------------
            workbook.Save("UpdatedAccent4Theme.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved as 'UpdatedAccent4Theme.xlsx' with new Accent4 color.");
        }
    }
}
