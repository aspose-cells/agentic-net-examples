// Title: Create a bold, centered header row that merges cells A‑D in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge cells A1 to D1 and set bold centered header | How to add a merged header row with bold text in an Excel workbook using Aspose.Cells | Aspose.Cells example for inserting a top row and applying style to merged cells in C# | Create a merged header spanning columns A-D with bold font using Aspose.Cells .NET | Set bold centered style on a merged range in Excel with Aspose.Cells C# code
// Tags: merged header row A-D Aspose.Cells | bold centered header style Aspose.Cells | insert top row Excel workbook Aspose.Cells | Excel header formatting C# Aspose.Cells | cell style for merged range Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// The sample creates a new workbook, inserts a row at the top, merges cells A1 through D1, writes a header text, defines a style with bold font and centered alignment, applies the style to the merged range, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a new row at the top (row index 0)
        sheet.Cells.InsertRows(0, 1);

        // Merge cells A1 to D1 (row 0, columns 0-3)
        sheet.Cells.Merge(0, 0, 1, 4);

        // Set the header text in the merged cell
        sheet.Cells[0, 0].PutValue("Header Title");

        // Create a style with bold font and centered alignment
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.HorizontalAlignment = TextAlignmentType.Center;

        // Apply the style to the merged range
        sheet.Cells[0, 0].SetStyle(headerStyle);

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
