// Title: Apply a dark gray accent style to a cell and verify its color using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set dark gray background for a specific cell | verify cell foreground color after applying a style with Aspose.Cells .NET | simulate a grayscale theme in an Excel workbook using Aspose.Cells | apply solid pattern style to a header cell in Aspose.Cells C# example | check if a cell style color matches the expected value in Aspose.Cells
// Tags: dark gray cell style Aspose.Cells | solid background pattern C# Excel | cell foreground color verification Aspose.Cells | grayscale theme simulation Aspose.Cells | C# workbook cell formatting Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, adds sample data, defines a dark gray solid background style to mimic a grayscale accent, applies the style to the header cell, reads back the cell's foreground color to confirm it matches the expected gray value, and saves the file as GrayscaleTheme.xlsx.
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

            // Fill sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(456);

            // NOTE: Theme manipulation APIs may not be available in all versions of Aspose.Cells.
            // The required grayscale accents are applied directly via cell style instead.

            // Create a style that uses a dark gray background (simulating Accent1)
            Style grayStyle = workbook.CreateStyle();
            grayStyle.ForegroundColor = Color.FromArgb(0xFF, 0x33, 0x33, 0x33); // Dark gray
            grayStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the header cell
            sheet.Cells["A1"].SetStyle(grayStyle);

            // Verify that the header cell displays the expected grayscale color
            Color headerBg = sheet.Cells["A1"].GetStyle().ForegroundColor;
            if (headerBg.ToArgb() == Color.FromArgb(0xFF, 0x33, 0x33, 0x33).ToArgb())
            {
                Console.WriteLine("Header cell displays correct grayscale accent.");
            }
            else
            {
                Console.WriteLine("Header cell color mismatch.");
            }

            // Define output file path
            string outputPath = "GrayscaleTheme.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
