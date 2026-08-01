// Title: Apply Theme Accent1 Fill to Imported Cells with Aspose.Cells for .NET
// Description: Creates a workbook, imports sample data into A2:A6, defines a solid style using the workbook's ThemeColor Accent1, applies the fill via a StyleFlag, and saves the result as DataWithAccent1Fill.xlsx.
// Keywords: Aspose.Cells ThemeColor Accent1 | C# cell fill style | apply theme color to range | StyleFlag cell shading | solid fill Aspose.Cells | Excel theme accent fill .NET
// Common Searches: Aspose.Cells set cell background to Accent1 | C# apply workbook theme color as fill | How to use StyleFlag for cell shading in Aspose.Cells | Apply solid theme fill to a range in .NET | ThemeColor fill example Aspose.Cells
// Developer Intent: Use the workbook’s Theme Accent1 color as the background fill for cells populated by an import routine.
// Use Cases: Visually highlight imported data with the document’s primary accent color. | Maintain consistent theming across generated reports without hard‑coding RGB values. | Create Excel files that automatically adapt when the workbook theme changes.
// AI Prompts: Show how to add a tint to the Accent1 fill color in the example. | Demonstrate applying Accent2 and Accent3 to separate imported ranges. | Explain how to retrieve the RGB value of the applied Accent1 theme color.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// Creates a workbook, imports sample data into A2:A6, defines a solid style using the workbook's ThemeColor Accent1, applies the fill via a StyleFlag, and saves the result as DataWithAccent1Fill.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ----- Data import simulation -----
            // Populate cells A2:A6 with sample data
            for (int i = 0; i < 5; i++)
            {
                cells[i + 1, 0].PutValue(i + 1); // Row index i+1 (A2..A6), column 0 (A)
            }

            // ----- Apply Theme Accent1 as fill color -----
            // Create a style with solid pattern and set the foreground theme color to Accent1
            Style accentStyle = workbook.CreateStyle();
            accentStyle.Pattern = BackgroundType.Solid;
            accentStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0); // No tint

            // Define a StyleFlag to apply cell shading (fill)
            StyleFlag flag = new StyleFlag();
            flag.CellShading = true;

            // Apply the style to the imported data range
            Aspose.Cells.Range importedRange = cells.CreateRange("A2:A6");
            importedRange.ApplyStyle(accentStyle, flag);

            // Save the workbook
            workbook.Save("DataWithAccent1Fill.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
