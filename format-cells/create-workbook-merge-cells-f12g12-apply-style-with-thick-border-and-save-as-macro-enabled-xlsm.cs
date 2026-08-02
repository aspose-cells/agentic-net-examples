// Title: C# – Merge cells F12:G12, apply a thick black border, and save as macro‑enabled XLSM using Aspose.Cells
// Description: Creates a new workbook, merges the range F12:G12, defines a style with thick black borders on all sides, applies the style to the merged cell, and saves the file as a macro‑enabled XLSM document with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# merge cells | thick border style Aspose.Cells | macro enabled XLSX XLSM save | cell formatting Aspose.Cells | C# Excel workbook creation
// Common Searches: how to merge cells F12:G12 with Aspose.Cells | apply thick black border to merged cells C# | save Aspose.Cells workbook as XLSM | Aspose.Cells style borders example | C# code for macro enabled Excel file
// Developer Intent: Generate a macro‑enabled XLSM workbook, merge F12:G12, style it with a thick black border, and persist the file.
// Use Cases: Designing report headers that span multiple columns with prominent borders. | Building macro‑enabled templates where styled merged cells serve as VBA entry points. | Automating invoice generation with bold title blocks that require thick borders.
// AI Prompts: Write C# code with Aspose.Cells to merge cells F12:G12, add a thick black border, and export as XLSM. | Explain how to create a border style and apply it to a merged cell range in Aspose.Cells for .NET. | Show the steps to keep a workbook macro‑enabled while applying cell formatting in C#.

using Aspose.Cells;
using System.Drawing;

// Creates a new workbook, merges the range F12:G12, defines a style with thick black borders on all sides, applies the style to the merged cell, and saves the file as a macro‑enabled XLSM document with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells F12:G12 (row 11, column 5, 1 row, 2 columns)
        cells.Merge(11, 5, 1, 2);

        // Create a style with thick borders on all sides
        Style borderStyle = workbook.CreateStyle();
        borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
        borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
        borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
        borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
        borderStyle.Borders[BorderType.TopBorder].Color = Color.Black;
        borderStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
        borderStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
        borderStyle.Borders[BorderType.RightBorder].Color = Color.Black;

        // Apply the style to the merged cell (upper‑left cell of the range)
        cells[11, 5].SetStyle(borderStyle);

        // Save the workbook as a macro‑enabled XLSM file
        workbook.Save("MergedCellWithBorder.xlsm", SaveFormat.Xlsm);
    }
}
