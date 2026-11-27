// Title: How to apply the theme’s Accent6 background color to the header row of a generated Excel report using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set header row background to theme Accent6 color | apply theme accent color to first row in Excel with Aspose.Cells .NET | create style with Accent6 background for header cells in Aspose.Cells | use StyleFlag to style header row with theme color in Aspose.Cells C# | set solid fill background for header row in generated workbook Aspose.Cells
// Tags: theme accent6 background styling Aspose.Cells | header row formatting C# Aspose.Cells | solid fill style application Aspose.Cells | styleflag usage for row styling Aspose.Cells | export workbook to xlsx Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example creates a new workbook, writes header titles to the first row, defines a solid‑fill style using the theme’s Accent6 (fallback LightBlue) color, applies the style to the entire header row with a StyleFlag, and saves the file as Report.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Define header titles
            string[] headers = { "ID", "Name", "Date", "Amount" };

            // Write headers to the first row (row index 0)
            for (int col = 0; col < headers.Length; col++)
            {
                sheet.Cells[0, col].PutValue(headers[col]);
            }

            // Use a predefined color for the header background (e.g., Accent6 equivalent)
            Color headerBackgroundColor = Color.LightBlue; // fallback color

            // Create a style that uses the selected background color
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = headerBackgroundColor;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the entire header row
            var headerRange = sheet.Cells.CreateRange(0, 0, 1, headers.Length);
            var styleFlag = new StyleFlag { All = true };
            headerRange.ApplyStyle(headerStyle, styleFlag);

            // Define output file path
            string outputPath = "Report.xlsx";

            // Save the workbook (lifecycle rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
