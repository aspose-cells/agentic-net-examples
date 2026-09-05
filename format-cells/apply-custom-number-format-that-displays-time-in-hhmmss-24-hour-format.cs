// Title: Apply a custom 24‑hour time format (hh:mm:ss) to a cell with Aspose.Cells for .NET
// AI Prompts: Write C# code that inserts the current system time into a worksheet cell and uses Aspose.Cells to set the cell's style to the custom pattern hh:mm:ss. | Show how to configure a Style object in Aspose.Cells to display a DateTime value as 24‑hour time in an Excel file. | Provide a snippet that formats a range of cells with the hh:mm:ss number format using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set cell number format to hh:mm:ss | How to display time in 24 hour format in Excel using Aspose.Cells | Custom number format for time values in Aspose.Cells .NET example | Formatting DateTime cells with Aspose.Cells style custom property | Save workbook with time formatted cells using Aspose.Cells for C#
// Tags: hh:mm:ss style Aspose.Cells | set cell custom format .NET Excel | time display format worksheet Aspose | Aspose.Cells style object number pattern | C# Excel time formatting Aspose.Cells

using Aspose.Cells;
using System;

// Creates a workbook, writes the current DateTime to cell A1, applies the custom number format "hh:mm:ss" to show 24‑hour time, and saves the file as TimeFormatted.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a time value into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(DateTime.Now);

        // Apply custom number format "hh:mm:ss"
        Style style = cell.GetStyle();
        style.Custom = "hh:mm:ss";
        cell.SetStyle(style);

        // Save the workbook (save rule)
        workbook.Save("TimeFormatted.xlsx");
    }
}
