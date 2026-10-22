// Title: How to merge cells B2 through D2 and center the text horizontally using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that merges the range B2:D2 on the first worksheet, inserts a value, applies horizontal center alignment, and saves the workbook as an .xlsx file. | Explain how to retrieve a cell's style, set its HorizontalAlignment property to Center, and reapply the style after merging cells using Aspose.Cells in a .NET project. | Provide a step‑by‑step example of merging a row of cells and formatting the merged cell's alignment with the Aspose.Cells API in C#.
// Common Searches: Aspose.Cells C# merge B2 D2 cells and center text horizontally | C# code to merge a range of cells and set alignment with Aspose.Cells | How to apply horizontal alignment to a merged cell using Aspose.Cells for .NET | Saving an Excel file after merging cells with Aspose.Cells in C#
// Tags: merge cell range B2:D2 Aspose.Cells C# | center text horizontally merged cell Aspose.Cells | apply cell style Aspose.Cells .NET | export workbook to xlsx Aspose.Cells

using Aspose.Cells;

// Creates a new workbook, merges cells B2 through D2 on the first worksheet, puts a value into the merged cell, centers the text horizontally via cell style, and saves the result as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Merge cells B2 through D2 (row index 1, column indices 1 to 3)
        sheet.Cells.Merge(1, 1, 1, 3);

        // Set a value for the merged cell
        sheet.Cells["B2"].PutValue("Merged Text");

        // Retrieve the style of the merged cell
        var style = sheet.Cells["B2"].GetStyle();

        // Center the text horizontally
        style.HorizontalAlignment = TextAlignmentType.Center;

        // Apply the updated style back to the cell
        sheet.Cells["B2"].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
