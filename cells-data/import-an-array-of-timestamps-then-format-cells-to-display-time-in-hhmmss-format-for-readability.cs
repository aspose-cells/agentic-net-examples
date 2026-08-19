// Title: C# – Import DateTime array into Excel and format as HH:mm:ss using Aspose.Cells
// Description: Creates a workbook, converts timestamp strings to DateTime objects, imports them vertically with Cells.ImportObjectArray, applies a custom "hh:mm:ss" number format via a style, and saves the result as TimestampsFormatted.xlsx.
// Keywords: Aspose.Cells C# import DateTime array | Excel time format HH:mm:ss | ImportObjectArray timestamp | custom number format Aspose.Cells | save workbook Aspose.Cells .NET
// Common Searches: Aspose.Cells import DateTime array C# | How to format Excel cells as HH:mm:ss with Aspose.Cells | Apply custom time style after importing timestamps .NET | Cells.ImportObjectArray time formatting example | Save Excel file with formatted timestamps Aspose
// Developer Intent: Load a collection of timestamps into a worksheet and display only the time component (HH:mm:ss) using Aspose.Cells for .NET.
// Use Cases: Generate a timesheet by importing raw datetime logs and showing only the time of day. | Export database datetime fields to Excel where the date is irrelevant for analysis. | Create a schedule workbook that lists start times extracted from logs with uniform formatting.
// AI Prompts: Give C# code that imports an array of DateTime objects into an Aspose.Cells worksheet and formats the cells with the HH:mm:ss custom number format. | Show how to apply a custom time style to a range after using Cells.ImportObjectArray in Aspose.Cells for .NET. | Explain how to ensure the output directory exists before saving a workbook that contains formatted timestamps.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsTimestampImport
{
    // Creates a workbook, converts timestamp strings to DateTime objects, imports them vertically with Cells.ImportObjectArray, applies a custom "hh:mm:ss" number format via a style, and saves the result as TimestampsFormatted.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample timestamps as strings
                string[] timestampStrings = new string[]
                {
                    "2023-01-01 08:15:30",
                    "2023-01-01 12:45:05",
                    "2023-01-01 16:20:55",
                    "2023-01-01 20:00:00"
                };

                // Convert the strings to DateTime objects and store them in an object array
                object[] timestampObjects = timestampStrings
                    .Select(s => DateTime.Parse(s))
                    .Cast<object>()
                    .ToArray();

                // Import the DateTime objects vertically starting at cell A1
                cells.ImportObjectArray(timestampObjects, 0, 0, true);

                // Create a style that displays time in HH:mm:ss format
                Style timeStyle = workbook.CreateStyle();
                timeStyle.Custom = "hh:mm:ss";

                // Apply the style to the imported range (column A)
                AsposeRange timeRange = cells.CreateRange(0, 0, timestampObjects.Length, 1);
                StyleFlag flag = new StyleFlag
                {
                    NumberFormat = true // Apply only the number format part of the style
                };
                timeRange.ApplyStyle(timeStyle, flag);

                // Define output file path
                string outputPath = "TimestampsFormatted.xlsx";

                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
