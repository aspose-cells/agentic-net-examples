// Title: Merge header row A1:D1 and apply navy blue fill using Aspose.Cells for .NET (C#)
// AI Prompts: Merge cells A1 through D1 on the first worksheet and set a solid navy background with Aspose.Cells in C#. | Create a navy‑filled style and apply it to a merged header range in a new workbook using the Aspose.Cells API. | Generate an XLSX file, merge the top row cells, apply navy fill, and save it as MergedHeader.xlsx with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# merge cells A1:D1 and set background color | how to apply navy fill to a merged header row using Aspose.Cells for .NET | C# example for merging first row cells and styling them with solid color in Aspose.Cells | set style for merged cells in an Aspose.Cells workbook
// Tags: merge cells range Aspose.Cells C# | solid navy fill style Aspose.Cells | header row formatting Aspose.Cells | create style with foreground color Aspose.Cells | save workbook as xlsx Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// Creates a new workbook, merges cells A1:D1 on the first worksheet, applies a solid navy background style to the merged header, and saves the file as MergedHeader.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the header range (e.g., A1:D1)
        int startRow = 0;      // Row 1 (zero‑based)
        int startColumn = 0;   // Column A
        int totalRows = 1;     // Only the header row
        int totalColumns = 4;  // Columns A to D

        // Merge the cells in the header range
        sheet.Cells.Merge(startRow, startColumn, totalRows, totalColumns);

        // Create a style with navy blue fill
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.Navy;
        style.Pattern = BackgroundType.Solid;

        // Apply the style to the merged cells (apply to the first cell of the range)
        StyleFlag flag = new StyleFlag();
        flag.All = true;
        sheet.Cells[startRow, startColumn].SetStyle(style, flag);

        // Save the workbook
        workbook.Save("MergedHeader.xlsx");
    }
}
