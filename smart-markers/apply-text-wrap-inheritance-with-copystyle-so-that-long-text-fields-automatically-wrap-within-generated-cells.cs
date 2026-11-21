// Title: How to copy text wrap formatting to other cells using SetStyle and GetStyle in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy wrap style from one cell to another | SetStyle GetStyle example for text wrapping in Aspose.Cells .NET | inherit text wrap formatting when generating Excel with Aspose.Cells | apply automatic text wrap to multiple cells using Aspose.Cells C# | copy cell style including IsTextWrapped property in Aspose.Cells
// Tags: copy cell style with SetStyle Aspose.Cells | text wrap formatting inheritance Aspose.Cells | SetStyle GetStyle text wrap C# | apply text wrap to generated Excel cells | Aspose.Cells style inheritance for long text

using Aspose.Cells;
using System;

// The example creates a workbook, defines a style with IsTextWrapped enabled, applies it to a source cell, and then copies that style to other cells using sourceCell.GetStyle() and SetStyle, ensuring long text automatically wraps in all generated cells before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style with text wrap enabled
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;

            // Apply the style to a source cell and set a long text value
            Cell sourceCell = sheet.Cells["A1"];
            sourceCell.PutValue("This is a very long text that should automatically wrap within the cell when the column width is limited.");
            sourceCell.SetStyle(wrapStyle);

            // Adjust column width to demonstrate wrapping
            sheet.Cells.SetColumnWidth(0, 20); // Column A

            // Copy the wrap style to other cells using SetStyle
            Cell targetCellB = sheet.Cells["B1"];
            targetCellB.PutValue("Another long text that needs wrapping.");
            targetCellB.SetStyle(sourceCell.GetStyle());

            Cell targetCellC = sheet.Cells["C1"];
            targetCellC.PutValue("Short text");
            targetCellC.SetStyle(sourceCell.GetStyle());

            // Save the workbook
            workbook.Save("WrappedTextDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
