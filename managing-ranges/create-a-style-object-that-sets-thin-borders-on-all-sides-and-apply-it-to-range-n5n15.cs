// Title: Apply Thin Black Borders to Range N5:N15 with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a Style with thin black borders on all four sides, configure a StyleFlag to affect only borders, and apply the style to the N5:N15 range in a workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells thin borders C# | apply style to range Aspose.Cells | C# set cell borders Aspose.Cells | StyleFlag borders Aspose.Cells | CreateStyle thin border Aspose.Cells
// Common Searches: Aspose.Cells set thin border N5 N15 | C# apply border style to range Aspose.Cells | How to use StyleFlag for borders in Aspose.Cells | Create and apply style to cells Aspose.Cells .NET | Apply borders to column N using Aspose.Cells
// Developer Intent: Create a Style with thin black borders on all sides and apply it to cells N5 through N15.
// Use Cases: Add uniform thin borders to a column of data for printable reports. | Visually separate a block of cells in a financial statement. | Prepare a table header range with consistent borders before populating data.
// AI Prompts: Generate C# code using Aspose.Cells to apply a red dashed border to range A1:C10. | Show how to modify the example to use a blue double‑line border instead of thin black. | Explain how to apply the same border style to multiple non‑contiguous ranges in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to create a Style with thin black borders on all four sides, configure a StyleFlag to affect only borders, and apply the style to the N5:N15 range in a workbook using Aspose.Cells for .NET.
class ApplyThinBorders
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the target range N5:N15 (use fully qualified Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("N5:N15");

            // Create a style and set thin black borders on all four sides
            Style style = workbook.CreateStyle();
            style.SetBorder(BorderType.LeftBorder,   CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder,  CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder,    CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

            // Configure the style flag to apply only border settings
            StyleFlag flag = new StyleFlag { Borders = true };

            // Apply the style to the specified range
            range.ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("StyledRange.xlsx");
            Console.WriteLine("Workbook saved successfully as StyledRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
