// Title: Merge header cells, apply bold font, and center text horizontally in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge A1:D1 and make the header bold and centered | How to format a merged header row with bold font and horizontal alignment using Aspose.Cells .NET | C# example for merging first row cells and applying bold centered style in an XLSX file with Aspose.Cells | Set bold style and center alignment on a merged cell in Aspose.Cells for .NET | Create a bold centered header in Excel using Aspose.Cells C# code
// Tags: merge cells A1:D1 Aspose.Cells | apply bold font Aspose.Cells C# | horizontal text alignment Aspose.Cells | header row styling Aspose.Cells | save workbook to XLSX Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, merges cells A1 through D1, inserts "Sales Report", applies a bold font and centers the text horizontally, then saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells A1 to D1 (first row)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns
        sheet.Cells.Merge(0, 0, 1, 4);

        // Set the header text in the merged cell
        Cell headerCell = sheet.Cells[0, 0];
        headerCell.PutValue("Sales Report");

        // Retrieve the current style of the cell
        Style style = headerCell.GetStyle();

        // Apply bold font
        style.Font.IsBold = true;

        // Center the text horizontally
        style.HorizontalAlignment = TextAlignmentType.Center;

        // Apply the modified style back to the cell
        headerCell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
