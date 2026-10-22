// Title: Apply italic font and light gray fill style to column Q using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply italic style to an entire column | Set light gray background for column Q in Excel using Aspose.Cells | How to use StyleFlag to format a whole column with Aspose.Cells .NET | Create and apply a custom style to column Q programmatically | Apply font and fill formatting to column Q with Aspose.Cells for .NET
// Tags: column-wide style application Aspose.Cells | italic font with solid gray fill Aspose.Cells | StyleFlag All attribute usage Aspose.Cells | formatting column Q Excel Aspose.Cells | custom style creation Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a workbook, defines a style with italic font and a solid light‑gray fill, sets a StyleFlag with All=true, and applies the style to column Q (index 16) via the Columns collection, then saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a new style
            Style style = workbook.CreateStyle();

            // Set the font to italic
            style.Font.IsItalic = true;

            // Set the fill to light gray
            style.ForegroundColor = Color.LightGray;
            style.Pattern = BackgroundType.Solid;

            // Define a style flag to apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Apply the style to the entire column Q (zero‑based index 16)
            // Use the Columns collection from the Cells object
            sheet.Cells.Columns[16].ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
