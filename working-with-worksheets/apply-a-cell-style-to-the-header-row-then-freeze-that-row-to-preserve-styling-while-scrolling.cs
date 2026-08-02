// Title: Style Header Row and Freeze It with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, writes header cells, defines a bold white font on a dark‑blue solid background, applies the style to the first row via ApplyRowStyle with a StyleFlag, freezes the top row at A2, and saves the file as HeaderStyledAndFrozen.xlsx.
// Keywords: Aspose.Cells C# ApplyRowStyle | Aspose.Cells FreezePanes | header row styling Aspose.Cells | freeze first row Excel C# | StyleFlag Aspose.Cells | Excel header formatting .NET
// Common Searches: Aspose.Cells style header row C# | How to freeze first row in Excel using Aspose.Cells | ApplyRowStyle example C# Aspose.Cells | FreezePanes at A2 Aspose.Cells .NET | Excel header formatting with Aspose.Cells
// Developer Intent: Apply custom formatting to the worksheet’s header row and keep it visible by freezing the pane.
// Use Cases: Generating reports where column titles need bold white text on a dark‑blue background that stays fixed while scrolling. | Exporting large data sets to Excel with a styled, frozen header to improve readability. | Automating workbook creation where multiple sheets require consistent header styling and frozen panes.
// AI Prompts: Write C# code using Aspose.Cells to apply a bold white font on a dark‑blue background to the first row and freeze that row at A2. | Show how to add borders and center alignment to the styled header while preserving the freeze pane in Aspose.Cells. | Provide an example that applies the same styled and frozen header to all worksheets in a workbook using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, writes header cells, defines a bold white font on a dark‑blue solid background, applies the style to the first row via ApplyRowStyle with a StyleFlag, freezes the top row at A2, and saves the file as HeaderStyledAndFrozen.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample header values
            cells[0, 0].PutValue("ID");
            cells[0, 1].PutValue("Name");
            cells[0, 2].PutValue("Score");

            // Create a style for the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;                     // Bold font
            headerStyle.Font.Color = Color.White;               // White font color
            headerStyle.ForegroundColor = Color.DarkBlue;       // Background color
            headerStyle.Pattern = BackgroundType.Solid;         // Solid fill

            // Define which style attributes to apply
            StyleFlag flag = new StyleFlag();
            flag.Font = true;          // Apply font settings
            flag.CellShading = true;   // Apply fill (background) settings

            // Apply the style to the first row (row index 0)
            cells.ApplyRowStyle(0, headerStyle, flag);

            // Freeze the first row so it stays visible while scrolling
            // Freeze at cell A2 with 1 frozen row and 0 frozen columns
            sheet.FreezePanes("A2", 1, 0);

            // Save the workbook
            workbook.Save("HeaderStyledAndFrozen.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
