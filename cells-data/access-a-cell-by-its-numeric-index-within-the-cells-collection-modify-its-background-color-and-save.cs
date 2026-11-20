// Title: How to access a cell by row and column indexes and set its background color with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set background color of a cell using row and column indexes | change cell fill color by numeric index Aspose.Cells .NET | apply solid yellow style to a specific cell in a workbook with Aspose.Cells | retrieve cell D6 using zero based indexes and style it in C# Aspose.Cells | save workbook after modifying cell style using Aspose.Cells API
// Tags: cell background styling Aspose.Cells | row column index cell access Aspose.Cells | solid fill pattern .NET | workbook save after styling Aspose.Cells | zero based cell indexing C#

using Aspose.Cells;
using System.Drawing;

// The example creates a new workbook, accesses cell D6 via zero‑based row (5) and column (3) indexes, optionally writes a value, applies a solid yellow background style, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Access a cell by its numeric row and column indexes (zero‑based)
        // Example: row index 5, column index 3 corresponds to cell D6
        Cell cell = sheet.Cells[5, 3];

        // (Optional) Set a value in the cell
        cell.PutValue("Sample");

        // Modify the cell's background color
        Style style = cell.GetStyle();
        style.ForegroundColor = Color.Yellow;      // Desired background color
        style.Pattern = BackgroundType.Solid;      // Apply solid fill
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
