// Title: How to apply a custom number format in Aspose.Cells for .NET to keep large integer IDs from showing scientific notation
// AI Prompts: Generate C# code that creates a workbook, writes an array of long values to a column, and applies a style with custom number format "0" to keep the full integer display. | Show how to create a Style object in Aspose.Cells, set its Custom property to "0", enable the Number format flag, and apply it to a specific range. | Demonstrate saving the workbook as an .xlsx file after formatting large numbers to avoid scientific notation.
// Common Searches: Aspose.Cells .NET prevent scientific notation for large numeric values | C# write big integer IDs to Excel without exponential format using Aspose.Cells | custom number format 0 in Aspose.Cells to display full integer | apply number format to a range in Aspose.Cells C# example | format large identifiers in Excel workbook with Aspose.Cells style flag
// Tags: integer display format Aspose.Cells | disable exponential notation Aspose.Cells | range style application Aspose.Cells C# | large identifier formatting Excel Aspose.Cells | custom integer format string Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a new workbook, writes an array of large long identifiers to column A, defines a style with the custom number format "0" to force full integer display, applies the style to the identifier range using a StyleFlag, and saves the file as LargeIdentifiers.xlsx.
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

            // Sample large identifiers that could be shown in scientific notation
            long[] ids = { 1234567890123L, 9876543210987L, 1234567890123456L };

            // Write the identifiers to column A
            for (int i = 0; i < ids.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(ids[i]);
            }

            // Create a style with a custom number format that forces full display (no scientific notation)
            Style style = workbook.CreateStyle();
            style.Custom = "0";               // Custom format: plain integer
            style.Number = 0;                 // Ensure the cell is treated as a number

            // Apply the style to the range containing the identifiers
            Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, ids.Length, 1);
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;         // Apply only the number format part of the style
            range.ApplyStyle(style, flag);

            // Save the workbook (lifecycle rule)
            workbook.Save("LargeIdentifiers.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
