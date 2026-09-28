// Title: How to merge cells A2:D2 into a single header and center its text with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to merge the range A2:D2, set a header value, and apply both horizontal and vertical center alignment. | Show how to create a style, apply it to a merged cell range, and save the workbook as an .xlsx file with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# merge A2:D2 and center text in the merged cell | C# code to create a header row by merging cells and aligning text with Aspose.Cells | How to apply centered style to a merged range in Aspose.Cells for .NET | Saving a workbook after merging cells and setting header value using Aspose.Cells C#
// Tags: merge cell range Aspose.Cells C# | center alignment style Aspose.Cells | header row creation Aspose.Cells .xlsx | apply style to merged cells Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// This example creates a new workbook, merges cells A2:D2 on the first worksheet, applies a style that centers the text horizontally and vertically, inserts a header title, and saves the file as MergedHeader.xlsx using Aspose.Cells for .NET.
class MergeHeaderExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range A2:D2 and merge it into a single cell
            AsposeRange headerRange = sheet.Cells.CreateRange("A2", "D2");
            headerRange.Merge();

            // Create a style to center the text horizontally and vertically
            Style centerStyle = workbook.CreateStyle();
            centerStyle.HorizontalAlignment = TextAlignmentType.Center;
            centerStyle.VerticalAlignment = TextAlignmentType.Center;

            // Apply the style to the merged range
            StyleFlag styleFlag = new StyleFlag
            {
                All = true // apply all style attributes
            };
            headerRange.ApplyStyle(centerStyle, styleFlag);

            // Optionally set a header text
            sheet.Cells["A2"].PutValue("Header Title");

            // Save the workbook (lifecycle save)
            workbook.Save("MergedHeader.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
