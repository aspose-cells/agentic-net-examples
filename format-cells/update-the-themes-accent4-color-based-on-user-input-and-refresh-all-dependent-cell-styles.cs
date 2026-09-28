// Title: Change the Accent4 theme color of an Excel workbook from a hex code and refresh all cell styles with Aspose.Cells for .NET
// AI Prompts: Set the workbook's ThemeColorType.Accent4 using a hex color string and force a refresh of all cells by reapplying their existing styles in C# with Aspose.Cells. | Generate C# code that loads an .xlsx file, converts a user‑provided hex value to System.Drawing.Color, updates the Accent4 theme color, iterates through every worksheet and cell to reapply styles, and saves the modified workbook.
// Common Searches: Aspose.Cells C# change Excel theme Accent4 color from hex value | how to refresh cell formatting after modifying theme colors in Aspose.Cells | set workbook theme accent color programmatically .NET | apply new theme color to existing Excel file and update dependent styles using Aspose.Cells | convert hex string to System.Drawing.Color for Aspose.Cells theme update
// Tags: update theme accent4 color Aspose.Cells | convert hex to System.Drawing.Color C# | reapply cell style to refresh theme colors | iterate all worksheets cells Aspose.Cells | save workbook after theme modification .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Loads an existing workbook, converts a user‑supplied hex string to a System.Drawing.Color, updates the Accent4 theme color via Workbook.SetThemeColor, iterates through every worksheet and cell to reapply each cell's style (forcing a refresh of theme‑based colors), and saves the result.
class UpdateThemeAccent4
{
    static void Main()
    {
        try
        {
            // User input for the new Accent4 color (hex format, e.g., "#FF5733")
            string userHexColor = "#FF5733";

            // Convert the hex string to a System.Drawing.Color
            Color newAccent4Color = ColorTranslator.FromHtml(userHexColor);

            // Load the existing workbook (replace with your actual file path)
            string inputFilePath = "InputWorkbook.xlsx";

            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException($"Input file not found: {inputFilePath}");

            Workbook workbook = new Workbook(inputFilePath);

            // Update the theme's Accent4 color using the Workbook API
            workbook.SetThemeColor(ThemeColorType.Accent4, newAccent4Color);

            // Refresh dependent cell styles by reapplying the existing style
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        Style style = cell.GetStyle();

                        // Reapply the same style to force a refresh of theme‑based colors
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            string outputFilePath = "OutputWorkbook.xlsx";
            workbook.Save(outputFilePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
