// Title: Set the workbook’s default style to the Light2 theme color for automatic row fill in Aspose.Cells for .NET (C#)
// Common Searches: how to make newly inserted rows use a theme color as background in Aspose.Cells C# | Aspose.Cells set default row fill to Light2 theme color | C# Aspose.Cells default style inheritance for inserted rows | apply workbook default style with solid fill using theme color in Aspose.Cells | set default background color for rows when inserting rows with Aspose.Cells .NET
// Tags: default row background Aspose.Cells | Workbook.DefaultStyle usage Aspose.Cells | insert rows inherit style Aspose.Cells | theme color solid fill Aspose.Cells | solid fill style with theme color .NET

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, retrieves the Light2 theme color, builds a solid‑fill style with that color, assigns it to Workbook.DefaultStyle, and demonstrates that any newly inserted rows automatically inherit the theme‑based background before saving the file.
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

            // Retrieve a theme color (use Accent1 which is available in all versions)
            Color themeColor = workbook.GetThemeColor(ThemeColorType.Accent1);

            // Create a style with a solid fill using the theme color
            Style themeStyle = workbook.CreateStyle();
            themeStyle.Pattern = BackgroundType.Solid;
            themeStyle.ForegroundColor = themeColor;

            // Set this style as the default style for the workbook.
            // Newly inserted rows will inherit this fill automatically.
            workbook.DefaultStyle = themeStyle;

            // Example: Insert a new row at index 5 (will use the default fill)
            sheet.Cells.InsertRows(5, 1);

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
