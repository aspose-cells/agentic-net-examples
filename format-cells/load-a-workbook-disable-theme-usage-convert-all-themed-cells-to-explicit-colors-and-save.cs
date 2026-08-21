// Title: Aspose.Cells .NET – Replace Excel Theme Colors with Direct RGB Values (C#)
// Description: A C# sample that loads an Excel workbook with Aspose.Cells, disables theme reliance, iterates every worksheet and cell, resolves each Font, Foreground, and Background ThemeColor to its actual RGB using Workbook.GetThemeColor, assigns the concrete Color, clears the ThemeColor reference, and writes the file so all colors are stored explicitly.
// Keywords: Aspose.Cells | C# | Excel theme color conversion | GetThemeColor | explicit RGB colors | remove ThemeColor reference | cell style manipulation | disable Excel theme | Workbook.GetThemeColor example | color conversion .NET
// Common Searches: Aspose.Cells convert theme colors to RGB C# | How to replace Excel theme colors with actual colors using Aspose.Cells | GetThemeColor usage example Aspose.Cells | Remove ThemeColor from workbook Aspose.Cells | Save Excel file with explicit colors Aspose
// Developer Intent: Swap every ThemeColor in a workbook’s cell styles for its exact RGB value and persist the changes.
// Use Cases: Ensure consistent appearance when the file is opened on machines lacking the original theme. | Prepare spreadsheets for PDF or image export where theme information is ignored. | Clean legacy workbooks before archiving to guarantee color fidelity across Excel versions.
// AI Prompts: Generate C# code with Aspose.Cells that iterates all cells and replaces Font, Foreground, and Background ThemeColor properties with the RGB colors returned by Workbook.GetThemeColor. | Explain the steps to disable theme usage in an Aspose.Cells workbook and convert all themed styles to explicit colors, noting any performance tips. | Provide a concise tutorial for clearing ThemeColor references in Aspose.Cells and saving the workbook with only concrete Color values.

using System;
using System.Drawing;
using Aspose.Cells;

// A C# sample that loads an Excel workbook with Aspose.Cells, disables theme reliance, iterates every worksheet and cell, resolves each Font, Foreground, and Background ThemeColor to its actual RGB using Workbook.GetThemeColor, assigns the concrete Color, clears the ThemeColor reference, and writes the file so all colors are stored explicitly.
class ConvertThemedCellsToExplicitColors
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all cells in the worksheet
            foreach (Cell cell in sheet.Cells)
            {
                // Get the current style of the cell
                Style style = cell.GetStyle();
                bool styleChanged = false;

                // Convert themed font color to explicit color
                if (style.Font.ThemeColor != null)
                {
                    // Resolve the actual theme color
                    Color explicitFontColor = workbook.GetThemeColor(style.Font.ThemeColor.ColorType);
                    // Apply explicit color
                    style.Font.Color = explicitFontColor;
                    // Remove theme reference
                    style.Font.ThemeColor = null;
                    styleChanged = true;
                }

                // Convert themed foreground (fill) color to explicit color
                if (style.ForegroundThemeColor != null)
                {
                    Color explicitFgColor = workbook.GetThemeColor(style.ForegroundThemeColor.ColorType);
                    style.ForegroundColor = explicitFgColor;
                    style.ForegroundThemeColor = null;
                    styleChanged = true;
                }

                // Convert themed background color to explicit color
                if (style.BackgroundThemeColor != null)
                {
                    Color explicitBgColor = workbook.GetThemeColor(style.BackgroundThemeColor.ColorType);
                    style.BackgroundColor = explicitBgColor;
                    style.BackgroundThemeColor = null;
                    styleChanged = true;
                }

                // Apply the modified style back to the cell if any changes were made
                if (styleChanged)
                {
                    cell.SetStyle(style);
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
