// Title: Highlight Error Cells with Theme Accent3 Background using Aspose.Cells for .NET (C#)
// Description: Learn how to detect cells that contain formula errors in a workbook and apply the workbook's Accent3 theme color as a solid background fill with Aspose.Cells for C#. The example creates error‑producing formulas, calculates them, builds a style using BackgroundThemeColor, scans the used range, and saves the highlighted file.
// Keywords: Aspose.Cells C# error cells | theme Accent3 background fill | highlight formula errors .NET | Cell.GetRichValue error detection | apply theme color to cells Aspose | conditional formatting error values | Aspose.Cells style BackgroundThemeColor | C# workbook error highlighting
// Common Searches: Aspose.Cells highlight error cells C# | apply theme Accent3 background to cells with errors | detect #DIV/0! cells using Aspose.Cells | set background theme color for error values .NET | C# code to color error cells in Excel with Aspose
// Developer Intent: Programmatically fill any cell that contains a formula error with the workbook’s Accent3 theme color.
// Use Cases: Automatically flag #DIV/0! or #NAME? results in financial models for quick visual review. | Create a reusable helper method that scans a worksheet, identifies error cells via Cell.GetRichValue().ErrorValue, and applies a solid Accent3 background. | Generate audit‑ready reports where error cells are visually distinguished using the document’s theme colors.
// AI Prompts: Write a C# function with Aspose.Cells that finds all error cells in a worksheet and sets their background to ThemeColorType.Accent3. | Show how to create a Style with BackgroundThemeColor = ThemeColor(ThemeColorType.Accent3) and apply it after workbook.CalculateFormula(). | Explain the steps to use Cell.GetRichValue().ErrorValue for error detection and then apply a solid fill using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// Learn how to detect cells that contain formula errors in a workbook and apply the workbook's Accent3 theme color as a solid background fill with Aspose.Cells for C#. The example creates error‑producing formulas, calculates them, builds a style using BackgroundThemeColor, scans the used range, and saves the highlighted file.
class ApplyAccent3ToErrorCells
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data that will generate errors
        cells["A1"].Formula = "=1/0";               // #DIV/0! error
        cells["B2"].Formula = "=UNKNOWNFUNC()";    // #NAME? error

        // Calculate formulas to materialize the errors
        try
        {
            workbook.CalculateFormula();
        }
        catch
        {
            // Expected exceptions for error cells; ignore
        }

        // Define a style that uses the theme's Accent3 color as background
        Style errorStyle = workbook.CreateStyle();
        errorStyle.Pattern = BackgroundType.Solid;
        errorStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent3, 0); // No tint

        // Scan all used cells and apply the style to those containing errors
        for (int row = 0; row <= cells.MaxDataRow; row++)
        {
            for (int col = 0; col <= cells.MaxDataColumn; col++)
            {
                Cell cell = cells[row, col];
                CellRichValue rich = cell.GetRichValue();

                // ErrorCellValueType enum default is 0 (no error). Non‑zero indicates an error.
                if (rich != null && rich.ErrorValue != 0)
                {
                    cell.SetStyle(errorStyle);
                }
            }
        }

        // Save the workbook
        workbook.Save("ErrorCellsAccent3.xlsx");
    }
}
