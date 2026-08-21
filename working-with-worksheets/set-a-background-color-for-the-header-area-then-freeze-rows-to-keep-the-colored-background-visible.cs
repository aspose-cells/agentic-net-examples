// Title: Color Header Row and Freeze It with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, writes headers to A1‑C1, defines a solid LightGray style using Style and StyleFlag, applies the style to the first row, freezes that row with FreezePanes, and saves the file as HeaderFreezeDemo.xlsx.
// Keywords: Aspose.Cells header background color C# | freeze first row Aspose.Cells | apply row style Aspose.Cells | Excel freeze panes C# | StyleFlag cell shading | solid fill header Aspose.Cells
// Common Searches: How to set header row background color in Excel using Aspose.Cells C# | Aspose.Cells C# freeze top row after styling | Apply solid fill to first row and keep it visible while scrolling | C# code to style and freeze header in Aspose.Cells workbook
// Developer Intent: Create a styled header row with a solid fill and keep it fixed by freezing the first row in an Excel file using Aspose.Cells for .NET.
// Use Cases: Generating reports where the header row is highlighted and remains visible during vertical scrolling. | Building spreadsheet templates with a colored, frozen header to improve readability of large data tables. | Exporting data sets where the first row serves as a persistent label column and must stay in view.
// AI Prompts: Write C# code with Aspose.Cells to apply a custom background color to the first row and freeze that row. | Show how to combine Style, StyleFlag, and FreezePanes in Aspose.Cells so a styled header stays visible while scrolling. | Provide an example that sets a gradient fill for the header row and freezes the top two rows in an Excel workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using System.Drawing;

// Creates a new workbook, writes headers to A1‑C1, defines a solid LightGray style using Style and StyleFlag, applies the style to the first row, freezes that row with FreezePanes, and saves the file as HeaderFreezeDemo.xlsx.
class SetHeaderBackgroundAndFreeze
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Fill header cells
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Score");

        // Create a style with a solid background color for the header row
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Pattern = BackgroundType.Solid;
        headerStyle.BackgroundColor = Color.LightGray; // desired background color

        // Define which style attributes to apply (background color)
        StyleFlag flag = new StyleFlag();
        flag.CellShading = true;

        // Apply the style to the entire first row (row index 0)
        cells.ApplyRowStyle(0, headerStyle, flag);

        // Freeze the first row so the colored header stays visible while scrolling
        worksheet.FreezePanes("A2", 1, 0); // freezes row 1 (the header)

        // Save the workbook
        workbook.Save("HeaderFreezeDemo.xlsx");
    }
}
