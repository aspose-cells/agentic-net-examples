// Title: Aspose.Cells C# – Apply Thin Black Borders to Range N5:N15
// Description: Demonstrates creating a Style with thin black borders on all sides, using a StyleFlag to limit the change to borders, defining the N5:N15 range, applying the style to that range, and saving the workbook as Output.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | thin border | cell borders | StyleFlag | apply style to range | N5:N15 | workbook formatting | .NET spreadsheet | border formatting
// Common Searches: Aspose.Cells set thin borders on a range | C# apply border style to cells N5 to N15 | How to use StyleFlag for borders only in Aspose.Cells | Create and apply a border style to a column range in .NET | Aspose.Cells example for formatting specific cells
// Developer Intent: Create a Style that adds thin black borders on every side of cells N5‑N15 and apply it without affecting other cell properties.
// Use Cases: Highlight a column of data in a financial report by surrounding each cell with a subtle border. | Prepare a printable invoice where the item list (N5‑N15) needs clear cell separation. | Standardize border formatting across multiple worksheets by reusing the same Style and StyleFlag.
// AI Prompts: Write C# code with Aspose.Cells to apply a dashed red border to range A1:C10 while preserving existing cell styles. | Show how to define a reusable Style that only sets left and right borders and apply it to several column ranges in a workbook. | Provide an example of applying different border colors to multiple non‑contiguous ranges in the same worksheet using StyleFlag.

using Aspose.Cells;
using System;
using System.Drawing;

// Demonstrates creating a Style with thin black borders on all sides, using a StyleFlag to limit the change to borders, defining the N5:N15 range, applying the style to that range, and saving the workbook as Output.xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style object
            Style style = workbook.CreateStyle();

            // Set thin black borders on all four sides
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

            // Prepare a StyleFlag to apply only border settings
            StyleFlag flag = new StyleFlag { Borders = true };

            // Define the target range N5:N15 (use fully qualified Aspose.Cells.Range)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("N5:N15");

            // Apply the style with the flag to the range
            range.ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
