// Title: How to merge cells B2:D2 and horizontally center text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that merges the range B2:D2 on the first worksheet and sets the horizontal alignment of the merged cell to Center. | Show how to apply a style to a merged cell in Aspose.Cells, including setting TextAlignmentType.Center for the merged range B2:D2. | Create a .NET example that merges a row of cells, writes a value into the merged area, and saves the workbook as an .xlsx file using Aspose.Cells.
// Common Searches: Aspose.Cells C# merge B2:D2 and center text horizontally | how to set horizontal alignment for merged cells in Aspose.Cells .NET | C# example merging cells and applying center alignment with Aspose.Cells library
// Tags: Aspose.Cells merge cell range C# | Aspose.Cells set horizontal alignment | center text in merged Excel cells .NET | apply style to merged range Aspose | save merged cells workbook as .xlsx

using Aspose.Cells;
using System;

// This C# example creates a new workbook, merges cells B2 through D2 on the first worksheet, centers the text horizontally in the merged cell, writes a sample value, and saves the file as MergedCells.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells B2 through D2 (row 1, columns 1 to 3)
        sheet.Cells.Merge(1, 1, 1, 3);

        // Get the style of the merged cell (top‑left cell B2)
        Style style = sheet.Cells["B2"].GetStyle();

        // Center the text horizontally
        style.HorizontalAlignment = TextAlignmentType.Center;

        // Apply the style back to the merged cell
        sheet.Cells["B2"].SetStyle(style);

        // Example text to show centering
        sheet.Cells["B2"].PutValue("Merged and Centered");

        // Save the workbook
        workbook.Save("MergedCells.xlsx");
    }
}
