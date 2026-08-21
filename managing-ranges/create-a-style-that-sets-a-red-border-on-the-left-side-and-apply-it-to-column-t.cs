// Title: Aspose.Cells .NET: Add a Red Left Border to Column T
// Description: Shows how to create a thin red left‑border style, limit the change with a StyleFlag, and apply it to column T (index 19) in an Excel workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# border style | red left border | column T formatting | StyleFlag | ApplyColumnStyle | Excel border color .NET | thin border | Excel column styling | Aspose.Cells example
// Common Searches: Aspose.Cells add red left border to column | C# apply border to specific column using StyleFlag | How to set left border color in Aspose.Cells | Apply style to column T in .NET | Set thin red border on Excel column with Aspose
// Developer Intent: Create a style with a thin red left border and apply it exclusively to column T.
// Use Cases: Visually separate a status column in financial reports. | Highlight a user‑input column in exported spreadsheets. | Maintain consistent left‑edge formatting across multiple workbooks.
// AI Prompts: Generate C# code with Aspose.Cells to apply a blue top border to column A. | Demonstrate applying different border styles to several columns using StyleFlag. | Explain how to reuse a single Style object for multiple columns without recreating it.

using System;
using System.Drawing;
using Aspose.Cells;

// Shows how to create a thin red left‑border style, limit the change with a StyleFlag, and apply it to column T (index 19) in an Excel workbook using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Create a style and set a red left border
        Style style = workbook.CreateStyle();
        style.Borders[BorderType.LeftBorder].Color = Color.Red;
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;

        // Create a style flag to apply only the left border
        StyleFlag flag = new StyleFlag();
        flag.LeftBorder = true;

        // Apply the style to column T (zero‑based index 19)
        cells.ApplyColumnStyle(19, style, flag);

        // Save the workbook
        workbook.Save("ColumnT_LeftBorder.xlsx");
    }
}
