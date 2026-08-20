// Title: C# – Merge Header Row (A1:D1) and Apply Navy Blue Fill with Aspose.Cells
// Description: Creates a new workbook, merges cells A1:D1, sets a solid navy background, optionally adds header text, and saves the file as MergedHeader.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | merge cells | header row | navy fill | solid background | Excel styling | range merge | Workbook | Worksheet
// Common Searches: Aspose.Cells merge header row C# | C# set cell background color navy Aspose.Cells | How to merge A1:D1 with Aspose.Cells | Apply solid fill to merged cells Aspose.Cells .NET
// Developer Intent: Combine the first‑row cells across multiple columns and style them with a navy blue background in an Excel workbook.
// Use Cases: Add a colored title bar to financial or sales reports. | Create a template where the top row serves as a branded section heading. | Generate dashboards with a navy‑blue header that spans several columns.
// AI Prompts: Generate C# Aspose.Cells code that merges A1:D1 and fills it with navy color. | Show how to create a solid navy style and apply it to a merged header range in a workbook using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a new workbook, merges cells A1:D1, sets a solid navy background, optionally adds header text, and saves the file as MergedHeader.xlsx using Aspose.Cells for .NET.
class MergeHeaderExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define a range that spans the header row (row 0) across columns A to D (0‑based indices 0‑3)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns
        int firstRow = 0;
        int firstColumn = 0;
        int totalRows = 1;      // only the header row
        int totalColumns = 4;   // columns A, B, C, D

        // Merge the defined range
        cells.Merge(firstRow, firstColumn, totalRows, totalColumns);

        // Create a style with navy blue fill
        Style style = workbook.CreateStyle();
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = Color.Navy;

        // Apply the style to the merged cell (top‑left cell of the range)
        cells[firstRow, firstColumn].SetStyle(style);

        // Optionally set a header value
        cells[firstRow, firstColumn].PutValue("Header");

        // Save the workbook
        workbook.Save("MergedHeader.xlsx");
    }
}
