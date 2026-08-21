// Title: Export Excel Theme Palette to a Worksheet with ARGB Values using Aspose.Cells for .NET
// Description: Creates a new worksheet named "ThemePalette", lists every ThemeColorType from a workbook, records each color's name, ARGB integer, and shows a styled sample cell, then auto‑fits columns and saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# export theme colors | Excel theme palette | ThemeColorType ARGB | sample color cell | .NET workbook theme extraction | save theme palette to sheet
// Common Searches: how to export Excel theme colors with Aspose.Cells | list ThemeColorType values in a worksheet C# | save workbook theme palette as XLSX | display ARGB values for Excel theme colors | create color sample cells in Aspose.Cells
// Developer Intent: Generate a worksheet that documents the workbook’s theme colors with their ARGB codes and visual samples for designers.
// Use Cases: Produce a design reference sheet showing each theme color name, ARGB code, and a colored preview. | Create a printable palette for brand guidelines by extracting theme colors from multiple workbooks. | Audit and compare default versus custom theme colors across different Excel files.
// AI Prompts: Write C# code with Aspose.Cells that iterates over ThemeColorType, writes the type name, ARGB value, and a styled sample cell to a new worksheet, then saves the workbook. | Modify the example to also include hexadecimal color codes and export accent colors in separate columns. | Generate a method that accepts a Workbook and returns a DataTable containing ThemeColorType, ARGB integer, hex string, and a base64 image of the color sample.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemePaletteExport
{
    // Creates a new worksheet named "ThemePalette", lists every ThemeColorType from a workbook, records each color's name, ARGB integer, and shows a styled sample cell, then auto‑fits columns and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Add a new worksheet to hold the theme palette information
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet sheet = workbook.Worksheets[sheetIndex];
            sheet.Name = "ThemePalette";

            // Write header row
            sheet.Cells["A1"].PutValue("Theme Color Type");
            sheet.Cells["B1"].PutValue("ARGB Value");
            sheet.Cells["C1"].PutValue("Sample");

            // Start writing data from the second row
            int currentRow = 1; // zero‑based index (row 2 in Excel)

            // Iterate through all ThemeColorType values (Background1 … FollowedHyperlink)
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve the actual color for the current theme type
                Color themeColor = workbook.GetThemeColor(type);

                // Write the theme type name
                sheet.Cells[currentRow, 0].PutValue(type.ToString());

                // Write the ARGB integer value (e.g., -16776961)
                sheet.Cells[currentRow, 1].PutValue(themeColor.ToArgb());

                // Create a style that uses the theme color as foreground (background) for visual sample
                Style sampleStyle = workbook.CreateStyle();
                sampleStyle.ForegroundColor = themeColor;
                sampleStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the sample cell and put a label
                Cell sampleCell = sheet.Cells[currentRow, 2];
                sampleCell.PutValue("Sample");
                sampleCell.SetStyle(sampleStyle);

                currentRow++;
            }

            // Auto‑fit columns for better readability
            sheet.AutoFitColumns();

            // Save the workbook to an XLSX file
            workbook.Save("ThemePaletteExport.xlsx", SaveFormat.Xlsx);
        }
    }
}
