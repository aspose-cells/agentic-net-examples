// Title: Create a custom percentage number format with a leading plus sign for positive values using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# custom number format to show + sign for positive percentages | How to format cells as +0.00% / -0.00% / 0.00% in Aspose.Cells | Apply a custom percentage format to a range of cells in a .NET workbook | Set positive percentage display with plus sign in Excel using Aspose.Cells API | C# Aspose.Cells style custom format string for percentages with sign
// Tags: custom percentage format Aspose.Cells | positive sign percentage style .NET | apply number format to cell range Aspose.Cells | custom number format string +0.00%;-0.00%;0.00% C# | style flag numberformat Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, writes fractional values to cells A1‑A3, defines a style with the custom format "+0.00%;-0.00%;0.00%" to display positive percentages with a leading plus sign, negative percentages with a minus sign, and zeros as 0.00%, applies this style to the specified range using a StyleFlag, and saves the file as CustomPercentageFormat.xlsx.
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

            // Sample data (values are fractions; they will be shown as percentages)
            sheet.Cells["A1"].PutValue(0.1234);   // 12.34%
            sheet.Cells["A2"].PutValue(-0.0567); // -5.67%
            sheet.Cells["A3"].PutValue(0);       // 0%

            // Create a style with a custom number format:
            // Positive values: +0.00%
            // Negative values: -0.00%
            // Zero values:      0.00%
            Style percentStyle = workbook.CreateStyle();
            percentStyle.Custom = "+0.00%;-0.00%;0.00%";

            // Apply the style to the range A1:A3
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:A3");
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true; // Apply only the number format part of the style
            range.ApplyStyle(percentStyle, flag);

            // Save the workbook
            workbook.Save("CustomPercentageFormat.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
