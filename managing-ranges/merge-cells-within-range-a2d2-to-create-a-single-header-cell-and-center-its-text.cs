// Title: Merge cells A2:D2 into a single header and center the text with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge A2:D2 and center header text | How to create a merged header row across columns A to D using Aspose.Cells | Applying horizontal and vertical alignment to a merged range in Aspose.Cells C# | Set merged cell style for header in Aspose.Cells workbook | Aspose.Cells example for merging cells and centering content in .NET
// Tags: cells.merge range Aspose.Cells C# | apply centered style to merged cells Aspose.Cells | header row spanning multiple columns Aspose.Cells | horizontal vertical alignment style Aspose.Cells | create merged header cell Aspose.Cells .NET

using System;
using Aspose.Cells;

// The sample creates a new workbook, merges cells A2 through D2 on the first worksheet, writes "Header Text" into the merged cell, defines a style with both horizontal and vertical center alignment, applies this style to the merged range, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Merge cells A2:D2 (row index 1, column index 0, 1 row, 4 columns)
            worksheet.Cells.Merge(1, 0, 1, 4);

            // Set header text in the merged cell
            worksheet.Cells["A2"].PutValue("Header Text");

            // Create a style that centers text horizontally and vertically
            var style = workbook.CreateStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;

            // Apply the style to the merged range using CreateRange
            var styleFlag = new StyleFlag { All = true };
            var range = worksheet.Cells.CreateRange("A2:D2");
            range.ApplyStyle(style, styleFlag);

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
