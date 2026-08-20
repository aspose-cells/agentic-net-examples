// Title: C# – Merge cells F12:G12, add thick black border, and save as macro‑enabled XLSM with Aspose.Cells
// Description: Shows how to create a new workbook using Aspose.Cells for .NET, merge the range F12:G12, apply a thick black border style to the merged cells, and export the result as a macro‑enabled XLSM file.
// Keywords: Aspose.Cells merge cells C# | thick border style Aspose.Cells | macro enabled XLSM save .NET | cell formatting Aspose.Cells | Excel workbook creation C# | apply borders to merged range | save workbook as XLSM | Aspose.Cells cell style API | C# Excel macro template
// Common Searches: Aspose.Cells merge F12:G12 C# | add thick border to merged cells Aspose.Cells | save workbook as macro enabled XLSM using .NET | C# code for merging cells and applying borders in Excel | how to create macro enabled file with Aspose.Cells
// Developer Intent: Create a workbook, merge a specific range, style it with a thick black border, and export it as a macro‑enabled XLSM file.
// Use Cases: Design a header that spans two columns in a financial report and highlight it with a bold border. | Build a macro‑enabled template where input sections are marked by merged cells with prominent borders for downstream VBA scripts. | Generate a printable invoice where the total amount cell merges two columns and is emphasized with a thick border.
// AI Prompts: Write C# code that merges cells F12:G12, applies a thick black border, and saves the workbook as an XLSM file using Aspose.Cells. | Explain how to create a reusable Style with thick borders in Aspose.Cells and apply it to multiple merged ranges before saving as a macro‑enabled workbook. | Show how to set border color and line style for all sides of a merged cell range in Aspose.Cells for .NET.

using System.Drawing;
using Aspose.Cells;

// Shows how to create a new workbook using Aspose.Cells for .NET, merge the range F12:G12, apply a thick black border style to the merged cells, and export the result as a macro‑enabled XLSM file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells F12:G12 (zero‑based row 11, column 5, 1 row, 2 columns)
        worksheet.Cells.Merge(11, 5, 1, 2);

        // Create a style with a thick border on all sides
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
        worksheet.Cells[11, 5].SetStyle(borderStyle);

        // Save the workbook as a macro‑enabled XLSM file
        workbook.Save("MergedWithBorder.xlsm", SaveFormat.Xlsm);
    }
}
