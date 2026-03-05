using System;
using System.Drawing;
using Aspose.Cells;

class FindFormulaCells
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a style to highlight formula cells (optional)
        Style highlightStyle = workbook.CreateStyle();
        highlightStyle.ForegroundColor = Color.Yellow;
        highlightStyle.Pattern = BackgroundType.Solid;

        // Iterate through all used cells in the worksheet
        foreach (Cell cell in cells)
        {
            // Check if the cell contains a formula
            if (cell.IsFormula)
            {
                // Output cell address and its formula
                Console.WriteLine($"Formula cell: {cell.Name}, Formula: {cell.Formula}");

                // Apply the highlight style to the formula cell
                cell.SetStyle(highlightStyle);
            }
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}