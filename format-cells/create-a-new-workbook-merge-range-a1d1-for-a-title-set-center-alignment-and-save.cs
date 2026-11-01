// Title: Create a new workbook, merge cells A1:D1 for a title, center-align the text, and save as XLSX with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge cells A1:D1 and center the title text | How to create a title row by merging cells and applying center alignment in Aspose.Cells | Save a workbook with a merged title row to XLSX using Aspose.Cells .NET | Set horizontal alignment for a merged cell range in Aspose.Cells C# | Initialize a workbook, merge a range, and apply style with Aspose.Cells
// Tags: merge range A1:D1 Aspose.Cells | center horizontal alignment merged cell C# | initialize workbook and save XLSX Aspose.Cells | apply style to merged cells Aspose.Cells | title row merge Aspose.Cells .NET

using Aspose.Cells;

// The example creates a new Workbook, merges cells A1 through D1 on the first worksheet, writes "Title" into the merged range, applies a centered horizontal alignment style, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Merge cells A1:D1 (row 0, column 0, 1 row, 4 columns)
        sheet.Cells.Merge(0, 0, 1, 4);

        // Set a title in the merged cell
        sheet.Cells["A1"].PutValue("Title");

        // Apply center alignment to the merged cell
        var style = sheet.Cells["A1"].GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        sheet.Cells["A1"].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
