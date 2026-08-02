// Title: Apply a Custom Theme Color to Outline Borders of a Range in the Second Worksheet (Aspose.Cells for .NET)
// Description: C# example that creates a workbook, adds or selects the second worksheet, defines the range B2:D5, builds a CellsColor using the Accent2 theme, applies a thin outline border with that theme color, and saves the file as CustomThemeBorderDemo.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells theme border color | C# set outline border theme | apply custom theme color to range borders | second worksheet border styling Aspose | CellsColor ThemeColorType Accent2 | thin outline border Aspose.Cells
// Common Searches: how to set a theme color for cell borders in Aspose.Cells | apply thin outline border with theme color to a range | border styling on second worksheet Aspose.Cells .NET | use ThemeColorType Accent2 for borders in C# | create CellsColor with theme in Aspose.Cells
// Developer Intent: Style the outline borders of a specific cell range on the second worksheet using a workbook theme color.
// Use Cases: Consistently brand tables on a secondary sheet by using the workbook’s Accent2 theme for borders. | Add a new worksheet to an existing workbook and immediately apply theme‑based borders to a data range before exporting. | Replace fixed RGB border colors with dynamic theme colors so the spreadsheet adapts to different Office themes.
// AI Prompts: Generate C# code with Aspose.Cells to apply a thin outline border using ThemeColorType.Accent3 to range A1:C3 on the third worksheet. | Show how to adjust the tint of a CellsColor theme border to a lighter shade and apply it to the inside borders of a range. | Explain how to read the current workbook theme and use those colors for borders on multiple ranges in Aspose.Cells.

using System;
using Aspose.Cells;
using System.Drawing;

// C# example that creates a workbook, adds or selects the second worksheet, defines the range B2:D5, builds a CellsColor using the Accent2 theme, applies a thin outline border with that theme color, and saves the file as CustomThemeBorderDemo.xlsx using Aspose.Cells.
class ApplyCustomThemeBorder
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Ensure there is a second worksheet (index 1)
            Worksheet worksheet;
            if (workbook.Worksheets.Count > 1)
            {
                worksheet = workbook.Worksheets[1];
            }
            else
            {
                // Add a new worksheet and get it as the second sheet
                int newIndex = workbook.Worksheets.Add();
                worksheet = workbook.Worksheets[newIndex];
            }

            // Define the range whose borders will be styled
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("B2:D5");

            // Create a CellsColor object and assign a theme color (e.g., Accent2)
            CellsColor themeBorderColor = workbook.CreateCellsColor();
            themeBorderColor.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0); // No tint

            // Apply the same thin border with the theme color to all outline edges of the range
            range.SetOutlineBorders(CellBorderType.Thin, themeBorderColor);

            // Save the workbook
            workbook.Save("CustomThemeBorderDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
