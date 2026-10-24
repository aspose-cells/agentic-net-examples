// Title: Add a thin red left border to column T in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to set a red left border on a specific column | How to apply a custom left border style to column T with Aspose.Cells .NET | Aspose.Cells style flag to format only borders of a column in C# | Create and apply a thin red left border to an entire column using Aspose.Cells | Apply border-only formatting to column T in an Excel file with Aspose.Cells C#
// Tags: Aspose.Cells set left border style C# | apply border-only style to column Aspose.Cells | thin red left border column formatting Aspose.Cells | StyleFlag border formatting Aspose.Cells .NET | column T styling Aspose.Cells workbook

using System;
using System.Drawing;
using Aspose.Cells;

// The sample creates a new workbook, defines a style with a thin red left border, uses a StyleFlag to limit formatting to borders only, applies the style to the entire column T (index 19), and saves the file as StyledColumnT.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Create a new style
        Style leftBorderStyle = workbook.CreateStyle();

        // Set left border to thin red line
        leftBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
        leftBorderStyle.Borders[BorderType.LeftBorder].Color = Color.Red;

        // Define a StyleFlag to apply only border formatting
        StyleFlag flag = new StyleFlag();
        flag.All = false;
        flag.Borders = true;

        // Column T is the 20th column (zero‑based index 19)
        int columnIndex = 19;

        // Apply the style to the entire column T
        sheet.Cells.Columns[columnIndex].ApplyStyle(leftBorderStyle, flag);

        // Save the workbook (lifecycle save)
        workbook.Save("StyledColumnT.xlsx");
    }
}
