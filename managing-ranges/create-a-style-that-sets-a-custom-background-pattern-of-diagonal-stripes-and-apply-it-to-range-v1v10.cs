// Title: How to create a diagonal stripe background style and apply it to cells V1:V10 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set diagonal stripe pattern for a specific cell range | Apply a custom background style to the range V1 V10 with Aspose.Cells .NET | Create an Excel style with foreground and background colors in C# using Aspose.Cells | Use StyleFlag.All to apply a style to multiple cells in Aspose.Cells | Save a workbook after styling cells with diagonal stripes in C#
// Tags: diagonal stripe background style Aspose.Cells | apply style to range V1 V10 C# | foreground background color settings Aspose.Cells | StyleFlag All usage Aspose.Cells | export styled workbook to Excel .NET

using Aspose.Cells;
using System;
using System.Drawing;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, defines a style with a diagonal stripe pattern (LightBlue foreground on White background), applies the style to cells V1 through V10 on the first worksheet, and saves the workbook as Output.xlsx.
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

            // Set diagonal stripe pattern
            style.Pattern = BackgroundType.DiagonalStripe;

            // Define foreground and background colors for the pattern
            style.ForegroundColor = Color.LightBlue;
            style.BackgroundColor = Color.White;

            // Apply the style to the range V1:V10
            AsposeRange range = sheet.Cells.CreateRange("V1", "V10");
            range.ApplyStyle(style, new StyleFlag() { All = true });

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
