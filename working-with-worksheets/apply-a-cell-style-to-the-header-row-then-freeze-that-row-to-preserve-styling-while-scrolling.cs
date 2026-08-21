// Title: Style the Header Row and Freeze the Top Row with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, writes column headings to the first row, defines a bold white font on a dark‑blue solid background, applies the style to the entire first row using a StyleFlag, freezes the pane at cell A2 so the header stays visible while scrolling, and saves the file as HeaderStyleAndFreeze.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel header style | ApplyRowStyle | StyleFlag | FreezePanes | freeze top row | worksheet formatting | Excel export template | bold header font | dark blue background
// Common Searches: Aspose.Cells style first row as header | How to freeze top row in Aspose.Cells C# | Apply custom style to header row Aspose.Cells | Freeze panes while preserving formatting Aspose.Cells | C# code to bold header and freeze pane in Excel
// Developer Intent: Apply a custom visual style to the first worksheet row and keep that row fixed during scrolling.
// Use Cases: Generating Excel reports with branded column headings that remain visible. | Creating reusable templates where the header row is automatically styled and frozen. | Exporting data grids to Excel with a clear, colored header for easier navigation.
// AI Prompts: Show C# code to style a header row with a dark background and freeze the top row using Aspose.Cells. | How can I apply different styles to multiple header rows and freeze them in Aspose.Cells for .NET? | Explain the interaction between StyleFlag and FreezePanes when preserving header formatting while scrolling.

using System;
using Aspose.Cells;
using System.Drawing;

// Creates a new workbook, writes column headings to the first row, defines a bold white font on a dark‑blue solid background, applies the style to the entire first row using a StyleFlag, freezes the pane at cell A2 so the header stays visible while scrolling, and saves the file as HeaderStyleAndFreeze.xlsx.
class HeaderStyleAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample header data
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Age");
        cells["C1"].PutValue("Country");

        // Define a style for the header row
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = Color.White;
        headerStyle.ForegroundColor = Color.DarkBlue;
        headerStyle.Pattern = BackgroundType.Solid;

        // Apply all style attributes
        StyleFlag flag = new StyleFlag();
        flag.All = true;

        // Apply the style to the first row (index 0)
        cells.ApplyRowStyle(0, headerStyle, flag);

        // Freeze the header row (first row) so it remains visible while scrolling
        // Freeze at cell A2 with 1 frozen row and 0 frozen columns
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("HeaderStyleAndFreeze.xlsx");
    }
}
