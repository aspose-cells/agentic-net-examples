// Title: How to apply a custom number format for 24‑hour time (hh:mm:ss) to a cell using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set custom number format to display time as hh:mm:ss | format DateTime cell as 24 hour time using Aspose.Cells .NET | apply custom time format hh:mm:ss to Excel cell with Aspose.Cells API | C# Aspose.Cells custom number format for time values | save workbook with time formatted as hh:mm:ss using Aspose.Cells
// Tags: custom number format hh:mm:ss Aspose.Cells | set cell time format .NET | 24‑hour time display Excel Aspose.Cells | C# Aspose.Cells time formatting | Excel workbook time format Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, writes a DateTime value (13:45:30) to cell A1, retrieves the cell's style, sets its Custom property to "hh:mm:ss" to show a 24‑hour time, reapplies the style, and saves the file as TimeFormatted.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a time value (13:45:30) into cell A1
        sheet.Cells["A1"].PutValue(new DateTime(2023, 1, 1, 13, 45, 30));

        // Retrieve the current style of the cell
        Style style = sheet.Cells["A1"].GetStyle();

        // Apply a custom number format to display time as "hh:mm:ss" (24‑hour format)
        style.Custom = "hh:mm:ss";

        // Set the modified style back to the cell
        sheet.Cells["A1"].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("TimeFormatted.xlsx");
    }
}
