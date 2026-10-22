// Title: How to apply a light gray solid background to rows 5 through 9 in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set solid fill color for rows 5 to 9 | apply background color to a range of rows in Excel with Aspose.Cells .NET | create row range 5-9 and style it using Aspose.Cells API | C# code to color multiple rows light gray in an Excel file with Aspose.Cells | range.ApplyStyle example for row range in Aspose.Cells
// Tags: apply style to row range Aspose.Cells | solid fill for Excel rows using Aspose.Cells | create range by row numbers Aspose.Cells | StyleFlag All property Aspose.Cells | background color rows Aspose.Cells .NET

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, defines a range covering rows 5 to 9, builds a style with a LightGray solid background, applies the style to the entire range using StyleFlag.All, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a range that covers rows 5 to 9 (address "5:9")
            // Use fully qualified name to avoid conflict with System.Range
            Aspose.Cells.Range range = sheet.Cells.CreateRange("5:9");

            // Create a style and set background color
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightGray;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the entire range
            StyleFlag flag = new StyleFlag { All = true };
            range.ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
