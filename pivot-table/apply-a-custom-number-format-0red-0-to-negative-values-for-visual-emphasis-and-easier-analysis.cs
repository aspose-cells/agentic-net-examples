// Title: Apply a custom number format '#,##0;[Red]-#,##0' to a cell range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom number format for negative values in Excel | How to display negative numbers in red with Aspose.Cells .NET | Apply '#,##0;[Red]-#,##0' format to a column using Aspose.Cells API | C# example for custom numeric formatting in an Aspose.Cells workbook | Using StyleFlag to apply number format to a range in Aspose.Cells
// Tags: Aspose.Cells numeric pattern C# | red negative value styling Aspose.Cells | StyleFlag number format flag | range.ApplyStyle custom style Aspose | Excel workbook generation .NET custom format

using System;
using Aspose.Cells;

// The sample creates a workbook, writes positive and negative numbers, defines a style with the custom format '#,##0;[Red]-#,##0' to show negatives in red, applies the style to the populated range using a StyleFlag that targets only the number format, and saves the file as 'CustomNumberFormat.xlsx'.
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

            // Sample data with positive and negative numbers
            double[] data = { 12500, -3400, 7200, -150 };
            for (int i = 0; i < data.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(data[i]);
            }

            // Create a style and set the custom number format
            // Format: positive numbers as normal, negative numbers in red with a minus sign
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0;[Red]-#,##0";

            // Apply the style to the range containing the numbers
            Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, data.Length, 1);
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true; // Apply only the number format part of the style
            range.ApplyStyle(customStyle, flag);

            // Save the workbook
            workbook.Save("CustomNumberFormat.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
