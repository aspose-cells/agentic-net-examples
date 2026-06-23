using System;
using Aspose.Cells;

class ApplyCustomTimeFormat
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a DateTime value that includes a time component
        // Example: 14:30:45 (2:30:45 PM)
        worksheet.Cells["A1"].PutValue(new DateTime(2023, 1, 1, 14, 30, 45));

        // Create a style object
        Style timeStyle = workbook.CreateStyle();

        // Apply a custom number format that displays time in 24‑hour "hh:mm:ss" format
        timeStyle.Custom = "hh:mm:ss";

        // Assign the style to the target cell
        worksheet.Cells["A1"].SetStyle(timeStyle);

        // Save the workbook to a file
        workbook.Save("CustomTimeFormat.xlsx");
    }
}