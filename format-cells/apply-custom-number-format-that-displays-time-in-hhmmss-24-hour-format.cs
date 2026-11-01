// Title: Apply a custom 24‑hour time format (hh:mm:ss) to a cell using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set cell number format to hh:mm:ss | display time as 24 hour format in Excel with Aspose.Cells .NET | apply custom time format to worksheet cell using Aspose.Cells API | C# Aspose.Cells format cell as time hh mm ss | save Excel file with 24‑hour time format using Aspose.Cells
// Tags: Aspose.Cells custom number format | time format hh:mm:ss Excel .NET | cell style time formatting Aspose | 24-hour time display workbook | C# Excel time formatting Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, inserts a DateTime value (13:45:30) into cell A1, defines a style with the custom number format "hh:mm:ss", applies the style to the cell, and saves the file as TimeFormat.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a time value (13:45:30) into cell A1
        // DateTime with only time component; date part is irrelevant for display
        DateTime timeValue = new DateTime(1, 1, 1, 13, 45, 30);
        sheet.Cells["A1"].PutValue(timeValue);

        // Create a style and set custom number format to "hh:mm:ss" (24‑hour format)
        Style timeStyle = workbook.CreateStyle();
        timeStyle.Custom = "hh:mm:ss";

        // Apply the style to the cell
        sheet.Cells["A1"].SetStyle(timeStyle);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("TimeFormat.xlsx");
    }
}
