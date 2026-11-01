// Title: Apply the workbook’s Accent3 theme color as a solid fill to a merged cell range using Aspose.Cells for .NET (C#)
// Common Searches: how to fill merged cells with a theme accent color using Aspose.Cells C# | Aspose.Cells C# apply workbook Accent3 color to merged range background | C# merge cells and set solid fill from workbook theme in Aspose.Cells | retrieve theme Accent3 color and use it for merged cells in Aspose.Cells .NET | set background color of merged cells to workbook theme color Aspose.Cells
// Tags: theme accent color fill merged cells | merged cells solid background Aspose.Cells | retrieve workbook theme color C# | set cell style foreground color Aspose.Cells | Aspose.Cells merged range styling

using Aspose.Cells;
using System.Drawing;

// The example creates a workbook, merges cells A1:C3, obtains the workbook's Accent3 theme color, applies it as a solid foreground fill to the merged range, and saves the file as MergedAccent3.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range to merge (e.g., A1:C3)
        int firstRow = 0;      // Row 1 (zero‑based)
        int firstColumn = 0;   // Column A
        int totalRows = 3;     // Span 3 rows (1‑3)
        int totalColumns = 3;  // Span 3 columns (A‑C)

        // Merge the cells across the specified rows and columns
        sheet.Cells.Merge(firstRow, firstColumn, totalRows, totalColumns);

        // Get the style of the top‑left cell of the merged range
        Style style = sheet.Cells[firstRow, firstColumn].GetStyle();

        // Retrieve the workbook's theme Accent3 color
        Color accent3 = workbook.GetThemeColor(ThemeColorType.Accent3);

        // Apply the Accent3 color to the fill of the merged cells
        style.ForegroundColor = accent3;
        style.Pattern = BackgroundType.Solid;

        // Assign the modified style back to the merged range (top‑left cell)
        sheet.Cells[firstRow, firstColumn].SetStyle(style);

        // Save the workbook (lifecycle rule)
        workbook.Save("MergedAccent3.xlsx");
    }
}
