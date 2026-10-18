// Title: How to apply a brand’s primary hex color to an Excel cell’s foreground using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set cell foreground color from HTML hex value | Convert hex color string to System.Drawing.Color for Excel styling with Aspose.Cells | Apply solid fill pattern with custom color to a cell using Aspose.Cells .NET | Write text and style a cell with brand color in a new workbook using Aspose.Cells | Save workbook after applying custom cell style with brand primary color in C#
// Tags: custom cell color Aspose.Cells | hex to System.Drawing.Color conversion C# | solid fill pattern Aspose.Cells style | brand-themed cell styling .NET | create workbook apply cell style Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, converts a brand's primary hex code to a System.Drawing.Color, writes text to cell A1, applies a style with the converted color as the foreground and a solid fill pattern, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Brand's primary color in hex
            string brandPrimaryHex = "#1A73E8"; // replace with actual brand hex code

            // Convert the hex color to a System.Drawing.Color
            Color brandPrimaryColor = ColorTranslator.FromHtml(brandPrimaryHex);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: update cell A1 with some text
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Brand Primary Color");

            // Apply a style that uses the brand's primary color for the foreground
            Style style = cell.GetStyle();
            style.ForegroundColor = brandPrimaryColor;          // set foreground color directly
            style.Pattern = BackgroundType.Solid;               // solid fill to show the color
            cell.SetStyle(style);

            // Save the workbook
            string outputPath = "BrandStyledWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
