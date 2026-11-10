// Title: Applying a yellow bold style to a union range G1:G3 and I1:I3 across all worksheets with WorksheetCollection.CreateUnionRange in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create union range G1:G3 I1:I3 on each worksheet | apply same style to multiple sheets using WorksheetCollection.CreateUnionRange | batch formatting cells G1:G3 and I1:I3 across all worksheets in Aspose.Cells | C# example of UnionRange.ApplyStyle with background color and bold font | how to iterate worksheets and style union ranges in Aspose.Cells .NET
// Tags: WorksheetCollection.CreateUnionRange collective styling | union range styling across worksheets | Aspose.Cells C# apply background color to union range | style union range G1:G3 I1:I3 | C# Aspose.Cells multiple sheet formatting

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a workbook with three worksheets, defines a yellow solid background and bold font style, then uses WorksheetCollection.CreateUnionRange to build a union range covering G1:G3 and I1:I3 on each sheet, applies the style to the range, and saves the file as UnionRangeFormatting.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Ensure there are multiple worksheets to demonstrate batch formatting
            Worksheet ws1 = workbook.Worksheets[0];
            ws1.Name = "Sheet1";
            Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
            Worksheet ws3 = workbook.Worksheets.Add("Sheet3");

            // Define a style to apply to the union range
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            style.Font.IsBold = true;

            // Apply the style to the union range on each worksheet
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Create a union range that includes cells G1:G3 and I1:I3 on the current worksheet
                UnionRange unionRange = workbook.Worksheets.CreateUnionRange("G1:G3,I1:I3", ws.Index);
                // Apply the style to the union range
                unionRange.ApplyStyle(style, new StyleFlag() { All = true });
            }

            // Save the workbook
            workbook.Save("UnionRangeFormatting.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
