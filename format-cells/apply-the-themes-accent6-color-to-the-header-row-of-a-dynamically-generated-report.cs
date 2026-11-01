// Title: Set the header row background to the workbook's Theme Accent6 color using Aspose.Cells for .NET
// Common Searches: how to use Aspose.Cells GetThemeColor to color a header row in C# | apply workbook theme Accent6 as solid fill to Excel header with Aspose.Cells .NET | C# Aspose.Cells style header row with theme color Accent6 | create solid background style from theme color in Aspose.Cells | set Excel header background to theme accent color using Aspose.Cells API
// Tags: apply theme accent6 background Aspose.Cells | header row solid fill style .NET | GetThemeColor API usage C# | StyleFlag range styling Aspose.Cells | create style with theme color Excel

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a new workbook, writes header titles to the first row, retrieves the workbook's Theme Accent6 color via GetThemeColor, builds a solid fill style with that color, applies the style to the entire header row using a StyleFlag, and saves the file as Report.xlsx.
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

            // Define header titles
            string[] headers = { "ID", "Name", "Date", "Amount" };

            // Write headers to the first row (row index 0)
            for (int col = 0; col < headers.Length; col++)
            {
                Cell cell = sheet.Cells[0, col];
                cell.PutValue(headers[col]);
            }

            // Retrieve the theme's Accent6 color using the correct API
            Color accent6 = workbook.GetThemeColor(ThemeColorType.Accent6);

            // Create a style that uses Accent6 as the background fill
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = accent6;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the entire header row
            StyleFlag flag = new StyleFlag { All = true };
            sheet.Cells.CreateRange(0, 0, 1, headers.Length).ApplyStyle(headerStyle, flag);

            // Save the workbook to a file
            workbook.Save("Report.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
