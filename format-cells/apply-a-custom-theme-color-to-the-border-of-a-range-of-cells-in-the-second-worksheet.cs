// Title: C# – Apply a custom theme color to range borders on the second worksheet with Aspose.Cells
// Description: Creates a workbook, adds a second worksheet if needed, sets the workbook Accent1 theme color to orange, builds a CellsColor that references this theme, selects range B2:D5 on the second sheet, and applies thin outline borders using the theme‑based color before saving the file.
// Keywords: Aspose.Cells C# theme color border | set workbook theme color Aspose | apply outline borders Aspose.Cells | second worksheet border formatting | CellsColor theme Aspose | custom theme border Excel .NET | range border color using theme | Aspose.Cells SetThemeColor example
// Common Searches: Aspose.Cells set custom theme color for borders | How to apply theme color to cell borders in C# | Add second worksheet and format borders with theme color | Create CellsColor from theme in Aspose.Cells | Outline border with Accent1 theme color Aspose | C# Aspose.Cells change workbook theme color
// Developer Intent: Apply a custom theme color to the borders of a specific cell range on the second worksheet using Aspose.Cells for .NET.
// Use Cases: Highlight a header block on a secondary sheet with orange borders that follow the workbook's Accent1 theme. | Build a multi‑sheet report template where all data tables use the same theme‑based outline borders for brand consistency. | Programmatically enforce theme‑aligned border styling across several worksheets in an automated Excel generation workflow.
// AI Prompts: Show how to adjust the tint of the theme border color to a lighter shade before applying it to the range. | Provide code to use thick and double border styles with the same theme color on multiple ranges in the second worksheet. | Explain how to revert to the default workbook theme after applying a custom theme color to specific borders.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, adds a second worksheet if needed, sets the workbook Accent1 theme color to orange, builds a CellsColor that references this theme, selects range B2:D5 on the second sheet, and applies thin outline borders using the theme‑based color before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (it contains one default worksheet)
            Workbook workbook = new Workbook();

            // Ensure there is a second worksheet; add one if necessary
            if (workbook.Worksheets.Count < 2)
                workbook.Worksheets.Add();

            // Get reference to the second worksheet (index 1)
            Worksheet secondSheet = workbook.Worksheets[1];

            // Define a custom theme color for Accent1
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.Orange);

            // Create a CellsColor object that refers to the theme color just set
            CellsColor themeBorderColor = workbook.CreateCellsColor();
            themeBorderColor.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0); // 0 tint

            // Define the range of cells whose borders will use the theme color
            AsposeRange targetRange = secondSheet.Cells.CreateRange("B2:D5");

            // Apply thin outline borders around the range using the theme‑based CellsColor
            targetRange.SetOutlineBorders(CellBorderType.Thin, themeBorderColor);

            // Save the workbook
            string outputPath = "SecondWorksheetThemeBorder.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
