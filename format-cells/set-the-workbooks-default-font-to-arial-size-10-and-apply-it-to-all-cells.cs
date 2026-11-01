// Title: How to set the default workbook font to Arial 10 and apply it to every cell using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set workbook default font to Arial size 10 | Apply a single style to all cells in an Aspose.Cells worksheet | How to change the default style and propagate it to the entire sheet with Aspose.Cells | Aspose.Cells default font for a new workbook in .NET | Use StyleFlag to apply default style to all cells in Aspose.Cells
// Tags: Aspose.Cells default style configuration C# | StyleFlag All usage Aspose.Cells | set workbook default font Arial Aspose.Cells | bulk apply style to worksheet cells Aspose.Cells | export workbook to Output.xlsx Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, modifies its DefaultStyle to use Arial size 10, uses a StyleFlag with All=true to apply that style to every cell in the first worksheet, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook's default font to Arial, size 10
        Style defaultStyle = workbook.DefaultStyle;
        defaultStyle.Font.Name = "Arial";
        defaultStyle.Font.Size = 10;
        workbook.DefaultStyle = defaultStyle; // optional, defaultStyle is a reference

        // Apply the default style to all cells in the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.ApplyStyle(defaultStyle, new StyleFlag() { All = true });

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
