using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class GradientFillAccentExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to which the gradient will be applied (e.g., A1:C5)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:C5");

            // Retrieve the theme colors Accent3 and Accent4 from the workbook's theme
            Color accent3 = workbook.GetThemeColor(ThemeColorType.Accent3);
            Color accent4 = workbook.GetThemeColor(ThemeColorType.Accent4);

            // Create a new style and set its fill to a two‑color gradient using the theme colors
            Style gradientStyle = workbook.CreateStyle();
            // GradientStyleType.Horizontal and variant 1 are used as an example
            gradientStyle.SetTwoColorGradient(accent3, accent4, GradientStyleType.Horizontal, 1);

            // Apply the style to the range (only cell shading is affected)
            StyleFlag flag = new StyleFlag { CellShading = true };
            range.ApplyStyle(gradientStyle, flag);

            // Adjust row height and column width for better visualisation (optional)
            sheet.Cells.SetRowHeight(0, 30);
            sheet.Cells.SetColumnWidth(0, 20);
            sheet.Cells.SetColumnWidth(1, 20);
            sheet.Cells.SetColumnWidth(2, 20);

            // Define output file path
            string outputPath = "GradientFillAccentDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}