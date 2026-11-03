// Title: Write a DateTime array to an Excel worksheet and format the cells as HH:mm:ss with Aspose.Cells for .NET
// Common Searches: how to write a DateTime[] to an Excel column using Aspose.Cells in C# | Aspose.Cells set cell format to display time as hh:mm:ss | apply a custom time number format to a range of cells in Aspose.Cells | C# import timestamps into Excel and format as HH:mm:ss with Aspose.Cells | use StyleFlag to apply number format to a cell range in Aspose.Cells
// Tags: Aspose.Cells write DateTime array to worksheet | time format style Aspose.Cells | apply StyleFlag number format Aspose.Cells | format Excel cells as HH:mm:ss C# | save workbook as Excel file Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, writes a predefined DateTime[] array into column A of the first worksheet, defines a Style with the custom number format "hh:mm:ss", applies this style to the populated range using a StyleFlag that enables number formatting, and saves the workbook as "Timestamps.xlsx" while handling any exceptions.
class TimestampFormatter
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample array of timestamps (DateTime values)
            DateTime[] timestamps = new DateTime[]
            {
                new DateTime(2023, 1, 1, 8, 15, 30),
                new DateTime(2023, 1, 1, 12, 45, 5),
                new DateTime(2023, 1, 1, 16, 0, 0),
                new DateTime(2023, 1, 1, 23, 59, 59)
            };

            // Write timestamps to column A starting from row 1 (0‑based index)
            for (int i = 0; i < timestamps.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(timestamps[i]);
            }

            // Create a style with the desired time format HH:mm:ss
            Style timeStyle = workbook.CreateStyle();
            timeStyle.Custom = "hh:mm:ss";

            // Apply the style to the range that contains the timestamps
            Aspose.Cells.Range timeRange = sheet.Cells.CreateRange(0, timestamps.Length, 0, 1);
            timeRange.ApplyStyle(timeStyle, new StyleFlag() { NumberFormat = true });

            // Save the workbook to a file
            workbook.Save("Timestamps.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
