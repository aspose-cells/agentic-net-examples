// Title: Write a DateTime[] to Excel and display values in HH:mm:ss format using Aspose.Cells for .NET
// AI Prompts: Create C# code that inserts a DateTime array into column A of a new workbook and applies a custom number format "hh:mm:ss" to those cells with Aspose.Cells. | Show how to define a Style with the time format and apply it to a specific range using StyleFlag in Aspose.Cells before saving the file. | Generate a .xlsx file named FormattedTimestamps.xlsx that contains the timestamps formatted as HH:mm:ss.
// Common Searches: Aspose.Cells C# format column as time hh:mm:ss | How to apply custom time format to a range of cells in Aspose.Cells | Write DateTime array to Excel using Aspose.Cells and keep time formatting | Saving timestamps with HH:mm:ss format in a .xlsx file via Aspose.Cells .NET
// Tags: Aspose.Cells apply custom time format | C# write DateTime array to worksheet | Aspose.Cells style range number format | export timestamps to Excel with HH:mm:ss | Aspose.Cells save workbook with time formatting

using System;
using Aspose.Cells;

// Creates a new workbook, writes a DateTime[] into column A, defines a style with the custom number format "hh:mm:ss", applies the style to the populated range, and saves the workbook as FormattedTimestamps.xlsx.
class TimestampFormatter
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample timestamps
            DateTime[] timestamps = new DateTime[]
            {
                new DateTime(2023, 1, 1, 8, 15, 30),
                new DateTime(2023, 1, 1, 12, 45, 5),
                new DateTime(2023, 1, 1, 16, 0, 0),
                new DateTime(2023, 1, 1, 23, 59, 59)
            };

            // Populate column A with the timestamps
            for (int i = 0; i < timestamps.Length; i++)
            {
                cells[i, 0].Value = timestamps[i];
            }

            // Define a style with the desired time format HH:mm:ss
            Style timeStyle = workbook.CreateStyle();
            timeStyle.Custom = "hh:mm:ss";

            // Apply the style to the range containing the timestamps
            Aspose.Cells.Range timeRange = cells.CreateRange(0, 0, timestamps.Length, 1);
            timeRange.ApplyStyle(timeStyle, new StyleFlag() { NumberFormat = true });

            // Save the workbook
            workbook.Save("FormattedTimestamps.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
