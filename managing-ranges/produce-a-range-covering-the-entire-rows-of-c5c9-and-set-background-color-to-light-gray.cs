// Title: Set a light gray background for rows 5‑9 by creating a row range with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create row range 5 to 9 and apply background color | How to apply a solid light gray fill to multiple rows using Aspose.Cells | Set background color for rows 5‑9 in an Excel file with Aspose.Cells .NET | Apply style to entire rows via CreateRange in Aspose.Cells C# | Styling rows 5 through 9 with a gray background using Aspose.Cells API
// Tags: row range styling Aspose.Cells C# | apply solid fill to rows Aspose.Cells | create range by row numbers Aspose.Cells | light gray background rows Aspose.Cells | Aspose.Cells style entire rows | set background color for specific rows Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// Creates a new workbook, defines a range covering rows 5 through 9, builds a style with a light‑gray solid background, applies the style to the entire range, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        // Create a range that covers the entire rows 5 to 9
        var range = worksheet.Cells.CreateRange("5:9");

        // Define a style with light gray background
        var style = workbook.CreateStyle();
        style.ForegroundColor = Color.LightGray;
        style.Pattern = BackgroundType.Solid;

        // Apply the style to the whole range
        var styleFlag = new StyleFlag();
        styleFlag.All = true;
        range.ApplyStyle(style, styleFlag);

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
