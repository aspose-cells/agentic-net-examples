// Title: Apply a grayscale accent theme to a workbook and validate cell background colors with Aspose.Cells for .NET
// Common Searches: how to set a custom gray accent color for cells using Aspose.Cells in C# | verify cell background color after applying a style with Aspose.Cells .NET | apply style to a range of cells with only cell shading using StyleFlag in Aspose.Cells | save workbook with a custom theme using Aspose.Cells for .NET | check if a cell's foreground color matches a specific RGB value in Aspose.Cells
// Tags: grayscale cell style Aspose.Cells | cell shading StyleFlag Aspose.Cells | foreground color verification Aspose.Cells | custom workbook theme Aspose.Cells .NET | save workbook with custom theme Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Retained for potential future drawing features
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// The example creates a new Workbook, defines a dark gray color as a custom accent, builds a solid fill Style, and applies it to the range A1:B2 using a StyleFlag that targets only cell shading. It then confirms that cell A1's foreground color matches the defined gray and saves the file as GrayscaleTheme.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Define a gray color to simulate Accent1
            Color accent1Gray = Color.FromArgb(64, 64, 64); // Dark gray

            // Create a style that uses the defined gray as the background fill
            Style grayStyle = workbook.CreateStyle();
            grayStyle.ForegroundColor = accent1Gray;
            grayStyle.Pattern = BackgroundType.Solid;

            // Define which style attributes to apply (only cell shading)
            StyleFlag flag = new StyleFlag { CellShading = true };

            // Apply the style to a range of cells
            AsposeRange range = sheet.Cells.CreateRange("A1:B2");
            range.ApplyStyle(grayStyle, flag);

            // Verify that a cell's background color matches the defined gray color
            Cell testCell = sheet.Cells["A1"];
            Color appliedColor = testCell.GetStyle().ForegroundColor;

            if (appliedColor.ToArgb() == accent1Gray.ToArgb())
            {
                Console.WriteLine("Theme applied correctly: cell background matches the defined gray color.");
            }
            else
            {
                Console.WriteLine("Theme verification failed.");
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("GrayscaleTheme.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
