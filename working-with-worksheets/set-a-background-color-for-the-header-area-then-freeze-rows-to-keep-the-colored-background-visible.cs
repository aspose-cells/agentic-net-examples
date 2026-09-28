// Title: Apply a light‑blue background to the header row and freeze the first row in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a solid light‑blue style, applies it to the first worksheet row with a StyleFlag, and freezes panes at cell A2 using Aspose.Cells. | Generate a complete Aspose.Cells example that sets a background color for header cells, uses StyleFlag to apply only cell shading, and keeps the header visible by freezing the top row.
// Common Searches: Aspose.Cells C# set header row background color and freeze top row | C# example to apply solid fill to Excel header and freeze panes with Aspose.Cells | How to keep colored header visible while scrolling in Excel using Aspose.Cells .NET | Freeze first row after styling header in Aspose.Cells workbook C#
// Tags: apply background color to header row Aspose.Cells | freeze first row Aspose.Cells C# | StyleFlag cell shading Aspose.Cells | FreezePanes A2 Aspose.Cells example | solid fill style Aspose.Cells workbook

using System;
using System.Drawing;
using Aspose.Cells;

// The program creates a new workbook, defines a solid light‑blue style for the first row using StyleFlag to apply cell shading, applies the style to the header row, freezes the top row at cell A2 so the colored header stays visible while scrolling, and saves the file as HeaderFreeze.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate header cells (example)
        worksheet.Cells["A1"].PutValue("Header 1");
        worksheet.Cells["B1"].PutValue("Header 2");
        worksheet.Cells["C1"].PutValue("Header 3");

        // Create a style for the header row
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Pattern = BackgroundType.Solid;          // Enable solid fill
        headerStyle.BackgroundColor = Color.LightBlue;       // Set background color

        // Define which style attributes to apply (background shading)
        StyleFlag flag = new StyleFlag();
        flag.CellShading = true;    // Apply background color

        // Apply the style to the entire first row (row index 0)
        worksheet.Cells.ApplyRowStyle(0, headerStyle, flag);

        // Freeze the first row so the colored header stays visible while scrolling
        // Freeze at cell A2 with 1 frozen row and 0 frozen columns
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("HeaderFreeze.xlsx");
    }
}
