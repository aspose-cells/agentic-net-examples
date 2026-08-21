// Title: C# – List Excel worksheets that use Light1 (Background1) theme background with Aspose.Cells
// Description: Load a workbook, scan each worksheet’s used cells, detect the Light1 (Background1) theme background via Style.BackgroundThemeColor, and output the names of worksheets that contain at least one such cell.
// Keywords: Aspose.Cells C# list worksheets | Light1 theme background Excel | Background1 theme color detection | ThemeColorType.Background1 Aspose | scan Excel cells for theme color | retrieve worksheet names by style
// Common Searches: how to find worksheets with Light1 background using Aspose.Cells | C# code to list sheets that contain Background1 themed cells | search Excel workbook for cells with theme background color | Aspose.Cells detect theme color in cells | list worksheets by cell style Aspose .NET
// Developer Intent: Return the names of all worksheets that contain at least one cell styled with the Light1 (Background1) theme background.
// Use Cases: Create an audit report of sheets that need theme‑color cleanup before publishing. | Target further processing (e.g., conditional formatting) only on sheets already using the Light1 background. | Automate documentation of worksheets that rely on the default Background1 theme color.
// AI Prompts: Generate C# code with Aspose.Cells that enumerates worksheet names containing any cell whose BackgroundThemeColor is ThemeColorType.Background1. | Explain a performance‑optimized method to locate Light1‑themed cells in large Excel files using Aspose.Cells. | Suggest an alternative technique to identify worksheets with Background1 theme cells without iterating every cell individually.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using System.Drawing;

// Load a workbook, scan each worksheet’s used cells, detect the Light1 (Background1) theme background via Style.BackgroundThemeColor, and output the names of worksheets that contain at least one such cell.
class ListWorksheetsWithLight1Background
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // List to hold worksheet names that contain cells with Light1 (Background1) theme background
        List<string> worksheetsWithLight1 = new List<string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            bool containsLight1 = false;

            // Iterate through all used cells in the worksheet
            foreach (Cell cell in worksheet.Cells)
            {
                // Get the style of the current cell
                Style style = cell.GetStyle();

                // Retrieve the background theme color (if any)
                ThemeColor bgTheme = style.BackgroundThemeColor;

                // Check if the background theme color is Background1 (Light1)
                if (bgTheme != null && bgTheme.ColorType == ThemeColorType.Background1)
                {
                    containsLight1 = true;
                    break; // No need to check further cells in this worksheet
                }
            }

            // If the worksheet contains at least one such cell, add its name to the list
            if (containsLight1)
            {
                worksheetsWithLight1.Add(worksheet.Name);
            }
        }

        // Output the names of worksheets that meet the criteria
        Console.WriteLine("Worksheets containing cells with Light1 (Background1) theme background:");
        foreach (string name in worksheetsWithLight1)
        {
            Console.WriteLine(name);
        }

        // Optionally, save the workbook if any modifications were made
        workbook.Save("output.xlsx");
    }
}
