// Title: Merge a header row range and apply navy blue fill color with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge cells A1:D1 and set background color | how to apply navy solid fill to merged header cells using Aspose.Cells | create range and merge header row in Aspose.Cells .NET example | set style foreground color navy for merged cells Aspose.Cells C#
// Tags: merge header row cells Aspose.Cells C# | apply navy solid fill style Aspose.Cells | create and merge range Aspose.Cells | set cell background color Aspose.Cells .NET | header formatting with Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, defines the range A1:D1 on the first worksheet, merges those cells, styles the merged header with a solid navy background, sets the header text, and saves the file as output.xlsx.
class MergeHeaderExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that spans the header row (e.g., A1:D1)
            string headerRange = "A1:D1";

            // Merge the cells in the defined range using a fully‑qualified Range type
            Aspose.Cells.Range range = sheet.Cells.CreateRange(headerRange);
            range.Merge();

            // Create a style for the merged header cell
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.Navy;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the top‑left cell of the merged range
            Cell headerCell = sheet.Cells["A1"];
            headerCell.SetStyle(headerStyle);

            // Set header text
            headerCell.PutValue("Header Title");

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
