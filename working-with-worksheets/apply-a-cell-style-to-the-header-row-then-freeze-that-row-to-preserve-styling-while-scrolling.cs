// Title: Style the header row with bold white text on dark blue and freeze it in an Excel sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a style with bold white font on a dark‑blue background, applies it to the first worksheet row, and freezes that row at cell A2. | Show how to use StyleFlag to copy all formatting attributes to a header row and then call FreezePanes to keep the header visible while scrolling. | Provide a complete Aspose.Cells example that formats the header row, freezes the pane, and saves the workbook as HeaderStyleAndFreeze.xlsx.
// Common Searches: Aspose.Cells C# how to apply a custom style to the first row and freeze the header | freeze top row after styling header with Aspose.Cells .NET | using StyleFlag to copy full row formatting and FreezePanes in Aspose.Cells
// Tags: header row styling with Aspose.Cells C# | freeze first row using FreezePanes Aspose.Cells | StyleFlag apply full row format Aspose.Cells | bold white font on dark blue background Excel | create and save workbook HeaderStyleAndFreeze.xlsx

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, defines a bold white‑on‑dark‑blue style, applies it to the first row, freezes that row at cell A2, and saves the file as HeaderStyleAndFreeze.xlsx.
class HeaderStyleAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add sample header values
        cells[0, 0].PutValue("ID");
        cells[0, 1].PutValue("Name");
        cells[0, 2].PutValue("Score");

        // Define a style for the header row
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = Color.White;
        headerStyle.ForegroundColor = Color.DarkBlue;
        headerStyle.Pattern = BackgroundType.Solid;

        // Create a flag to apply all style attributes
        StyleFlag flag = new StyleFlag { All = true };

        // Apply the style to the first row (row index 0)
        cells.ApplyRowStyle(0, headerStyle, flag);

        // Freeze the header row so it remains visible while scrolling
        // Freeze at cell A2 (row index 1) with 1 frozen row and 0 frozen columns
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("HeaderStyleAndFreeze.xlsx");
    }
}
