// Title: Apply a custom number format '#,##0;[Red]-#,##0' to highlight negative values in Excel using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom number format to display negative numbers in red | How to format negative Excel values with red color using Aspose.Cells .NET | C# example applying '#,##0;[Red]-#,##0' format with Aspose.Cells | Apply StyleFlag NumberFormat for red negative numbers in Aspose.Cells workbook | Create and apply custom number format to a cell range in Aspose.Cells C#
// Tags: custom number format Aspose.Cells C# | red negative number style Aspose.Cells | StyleFlag NumberFormat usage Aspose.Cells | range.ApplyStyle custom format .NET | Excel negative values formatting Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a workbook, fills cells A1‑A5 with positive and negative numbers, defines a style with the custom number format '#,##0;[Red]-#,##0' to show negatives in red, applies the style to the range via a StyleFlag, and saves the file as CustomNumberFormat.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (positive and negative values)
            sheet.Cells["A1"].PutValue(12345);
            sheet.Cells["A2"].PutValue(-6789);
            sheet.Cells["A3"].PutValue(0);
            sheet.Cells["A4"].PutValue(2500);
            sheet.Cells["A5"].PutValue(-1500);

            // Create a style object
            Style customStyle = workbook.CreateStyle();

            // Apply the custom number format: positive numbers as normal,
            // negative numbers in red with a minus sign
            customStyle.Custom = "#,##0;[Red]-#,##0";

            // Create a style flag to indicate which style attributes are applied
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true; // enable custom number format

            // Apply the style to the range containing the data
            Aspose.Cells.Range dataRange = sheet.Cells.CreateRange("A1:A5");
            dataRange.ApplyStyle(customStyle, flag);

            // Save the workbook (lifecycle save rule)
            workbook.Save("CustomNumberFormat.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
