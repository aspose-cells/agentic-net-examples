// Title: Set a cell's background color using numeric row/column indexes with Aspose.Cells for .NET
// Description: Creates a workbook, accesses a cell via zero‑based row and column indexes (cells[2,3]), assigns a value, applies a solid yellow background style, and saves the file as ModifiedCellBackground.xlsx.
// Keywords: Aspose.Cells C# background color | access cell by index Aspose.Cells | apply solid fill style .NET | save workbook after styling | numeric row column cell reference
// Common Searches: Aspose.Cells set cell color by row and column numbers | C# change background of a specific cell using Aspose.Cells | how to style a cell with solid fill in Aspose.Cells .NET | save workbook after applying cell style Aspose
// Developer Intent: Apply a solid background color to a cell identified by its numeric row and column indexes and persist the workbook.
// Use Cases: Highlight header rows in generated reports by coloring cells via index positions. | Mark error cells in automated spreadsheets with a red background using row/column coordinates. | Prepare template placeholders with distinct colors before distributing the workbook.
// AI Prompts: Write C# code that uses Aspose.Cells to set a blue background on the cell at row 5, column 2 and save as Report.xlsx. | Show how to loop through rows 0‑9 and columns 0‑4, applying a light gray solid fill to each cell with Aspose.Cells for .NET. | Explain how to create a reusable Style object for a green background and apply it to multiple cells accessed by numeric indexes.

using System.Drawing;
using Aspose.Cells;

// Creates a workbook, accesses a cell via zero‑based row and column indexes (cells[2,3]), assigns a value, applies a solid yellow background style, and saves the file as ModifiedCellBackground.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Access a cell by numeric row and column indexes (zero‑based)
        // Example: row 2, column 3 corresponds to cell D3 in Excel notation
        Cell targetCell = cells[2, 3];
        targetCell.PutValue("Background Demo");

        // Create a style, set its background color, and apply it to the cell
        Style bgStyle = workbook.CreateStyle();
        bgStyle.BackgroundColor = Color.Yellow;      // Desired background color
        bgStyle.Pattern = BackgroundType.Solid;      // Ensure the color is visible
        targetCell.SetStyle(bgStyle);

        // Save the workbook to a file
        workbook.Save("ModifiedCellBackground.xlsx");
    }
}
