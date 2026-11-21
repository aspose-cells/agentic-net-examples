// Title: How to merge cells B2‑D2 and center the text horizontally with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge B2:D2 and align text to center | C# code to merge a range of cells and set horizontal alignment using Aspose.Cells | center text in a merged Excel range with Aspose.Cells .NET library | merge cells and apply center alignment in an Aspose.Cells workbook example | programmatically merge cells and style horizontal alignment in Aspose.Cells C#
// Tags: merge cell range Aspose.Cells .NET | horizontal center alignment style Aspose.Cells | set cell value in merged range C# | Aspose.Cells workbook save to XLSX | apply style to specific cells Aspose.Cells

using Aspose.Cells;

// The sample creates a new workbook, merges cells B2 through D2 on the first worksheet, inserts text into the merged cell, defines a style with horizontal center alignment, applies the style to the merged range, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Merge cells B2 through D2 (zero‑based indexes: column 1 to 3, row 1)
        sheet.Cells.Merge(1, 1, 1, 3);

        // Optional: put some text into the merged cell
        sheet.Cells["B2"].PutValue("Merged Text");

        // Create a style with horizontal center alignment
        var style = workbook.CreateStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;

        // Define which style attributes to apply
        var styleFlag = new StyleFlag();
        styleFlag.HorizontalAlignment = true;

        // Apply the style to the merged cell
        sheet.Cells["B2"].SetStyle(style, styleFlag);

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
