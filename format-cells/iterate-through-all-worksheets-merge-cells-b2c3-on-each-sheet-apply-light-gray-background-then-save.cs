// Title: C# – Merge B2:C3 on Every Worksheet and Apply Light Gray Fill with Aspose.Cells
// Description: Creates a workbook, adds an extra sheet, iterates through all worksheets, merges the range B2:C3, applies a solid light‑gray background, and saves the result as MergedCellsLightGray.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | merge cells | B2:C3 | background color | light gray fill | iterate worksheets | style merged range | Excel automation | format cells
// Common Searches: Aspose.Cells merge cells on all sheets | C# set background color for merged range | how to style merged cells Aspose.Cells | loop through worksheets and format cells .NET | merge B2:C3 and fill gray Aspose.Cells
// Developer Intent: Merge B2:C3 on each worksheet, color it light gray, and save the workbook.
// Use Cases: Create a consistent header block across multiple sheets. | Highlight a summary area on every worksheet with a shaded merged cell. | Prepare a template where the title region (B2:C3) is uniformly styled.
// AI Prompts: Write C# code with Aspose.Cells that merges B2:C3 on every worksheet and sets a light gray background. | Show how to loop through all worksheets in a workbook and apply a solid gray style to a merged range. | Explain best practices for reusing a Style object when formatting merged cells across multiple sheets in Aspose.Cells.

using Aspose.Cells;
using System.Drawing;

// Creates a workbook, adds an extra sheet, iterates through all worksheets, merges the range B2:C3, applies a solid light‑gray background, and saves the result as MergedCellsLightGray.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (contains one default worksheet)
        Workbook workbook = new Workbook();

        // Add an extra worksheet so we have more than one sheet to process
        workbook.Worksheets.Add();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Merge cells B2:C3.
            // B2 is row 1, column 1 (zero‑based). The range spans 2 rows and 2 columns.
            sheet.Cells.Merge(1, 1, 2, 2);

            // Prepare a style with a light gray background
            Style style = sheet.Cells[1, 1].GetStyle();
            style.ForegroundColor = Color.LightGray;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the merged cell (top‑left cell of the merged area)
            sheet.Cells[1, 1].SetStyle(style);
        }

        // Save the workbook to a file
        workbook.Save("MergedCellsLightGray.xlsx");
    }
}
