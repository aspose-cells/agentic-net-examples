// Title: Apply a solid yellow background to a cell using numeric row/column indexes with Aspose.Cells for .NET and save as XLSX
// AI Prompts: Create a yellow solid fill style, assign it to the cell at row index 2 and column index 3 in a worksheet, and save the workbook to 'ModifiedCellBackground.xlsx' using Aspose.Cells for C#. | Retrieve a cell from the Cells collection by its zero‑based row and column numbers, set its value, apply a background color style, and export the workbook with Aspose.Cells .NET.
// Common Searches: Aspose.Cells C# set background color of a cell using row and column indexes | How to style a specific cell by numeric indexes in Aspose.Cells .NET | Saving a workbook after applying a solid fill style to a cell with Aspose.Cells | Zero‑based cell indexing example Aspose.Cells C# background color
// Tags: background color style Aspose.Cells C# | numeric cell indexing Aspose.Cells | apply solid fill Aspose.Cells | export workbook xlsx Aspose.Cells | zero‑based cell access Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The program creates a new workbook, accesses cell D3 via zero‑based row and column indexes, sets a value, applies a solid yellow background style, and saves the file as ModifiedCellBackground.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the Cells collection
        Cells cells = worksheet.Cells;

        // Access a cell by numeric row and column indexes (zero‑based)
        // Example: row 2, column 3 corresponds to cell D3 in Excel notation
        Cell targetCell = cells[2, 3];
        targetCell.PutValue("Demo");

        // Create a style and set its background color
        Style bgStyle = workbook.CreateStyle();
        bgStyle.BackgroundColor = Color.Yellow;      // Desired background color
        bgStyle.Pattern = BackgroundType.Solid;      // Ensure the color is visible

        // Apply the style to the selected cell
        targetCell.SetStyle(bgStyle);

        // Save the workbook (lifecycle rule)
        workbook.Save("ModifiedCellBackground.xlsx");
    }
}
