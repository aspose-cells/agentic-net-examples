// Title: How to color Excel header cells with a custom RGB fill using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a style with a solid background using Color.FromArgb(100,150,200) and applies it to cells A1:D1 in an Aspose.Cells workbook. | Generate a snippet that sets a custom RGB fill for the first row of a worksheet and saves the file as HeaderColored.xlsx using Aspose.Cells.
// Common Searches: Aspose.Cells C# set RGB background color for first row of worksheet | how to apply custom fill color to header cells in Excel using Aspose.Cells .NET | C# Aspose.Cells change header row background to specific color | using Color.FromArgb with Aspose.Cells to style cells
// Tags: set cell fill color Aspose.Cells C# | RGB background style for Excel header Aspose.Cells | apply solid fill to range Aspose.Cells | save workbook with colored header Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// The example creates a new workbook, defines a style with a solid RGB fill (100,150,200), applies this style to cells A1 through D1 as header cells, adds sample header text, and saves the workbook as HeaderColored.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define header range (first row, columns A to D)
        int headerRow = 0;          // Row index for header (0‑based)
        int startColumn = 0;       // Column A
        int endColumn = 3;         // Column D

        // Create a style with a custom fill color using RGB values (e.g., 100,150,200)
        Style headerStyle = workbook.CreateStyle();
        headerStyle.ForegroundColor = Color.FromArgb(100, 150, 200);
        headerStyle.Pattern = BackgroundType.Solid;

        // Apply the style to each header cell and set a sample value
        for (int col = startColumn; col <= endColumn; col++)
        {
            Cell cell = sheet.Cells[headerRow, col];
            cell.PutValue($"Header {col + 1}");
            cell.SetStyle(headerStyle);
        }

        // Save the workbook to a file
        workbook.Save("HeaderColored.xlsx");
    }
}
