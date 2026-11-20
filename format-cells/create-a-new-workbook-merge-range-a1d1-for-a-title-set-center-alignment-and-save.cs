// Title: Create a new workbook, merge cells A1:D1 into a centered title row, and save as XLSX with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge A1:D1 and center text in a workbook | How to set horizontal alignment for a merged cell using Aspose.Cells .NET | Saving a workbook with a merged title row using Aspose.Cells in C# | Create and style a title row by merging cells with Aspose.Cells for .NET | C# code to merge cells and apply center alignment in Aspose.Cells
// Tags: centered merged cell style Aspose.Cells | workbook creation and save Aspose.Cells .NET | title row merge formatting Aspose.Cells | apply style to merged range Aspose.Cells C# | cell merging with Aspose.Cells .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates how to create a new Workbook, merge the range A1:D1 on the first worksheet, insert the text "Title" into the merged cells, apply a style that centers the text horizontally, and save the file as MergedTitle.xlsx using Aspose.Cells for .NET in C#.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Merge cells A1:D1 (row 0, column 0, 1 row, 4 columns)
            cells.Merge(0, 0, 1, 4);

            // Set the title text in the merged cell
            cells["A1"].PutValue("Title");

            // Create a style with center horizontal alignment
            Style style = workbook.CreateStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;

            // Apply the style to the merged cell
            cells["A1"].SetStyle(style);

            // Save the workbook to a file
            workbook.Save("MergedTitle.xlsx");
        }
    }
}
