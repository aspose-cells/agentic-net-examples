// Title: Aspose.Cells for .NET – Enable 1904 Date System and Write Legacy Dates (C#)
// Description: Create a new Workbook, activate the 1904 date system via Workbook.Settings.Date1904, convert a .NET DateTime to an Excel serial number with CellsHelper.GetDoubleFromDateTime, write the value to a cell, apply a built‑in date format, and save the file—ensuring compatibility with legacy Mac Excel date calculations.
// Keywords: Aspose.Cells | C# | .NET | 1904 date system | Workbook.Settings.Date1904 | CellsHelper.GetDoubleFromDateTime | Excel serial date | legacy Mac Excel | date formatting | save workbook
// Common Searches: Aspose.Cells enable 1904 date system | How to set 1904 date system in C# | Convert DateTime to Excel serial number 1904 Aspose | Legacy Mac Excel date compatibility Aspose.Cells | Set workbook date system to 1904 using Aspose | Apply date format after 1904 conversion Aspose.Cells
// Developer Intent: Turn on the 1904 date system and write dates that respect it in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate workbooks compatible with older Mac Excel files that use the 1904 date base. | Create serial dates for legacy formulas that expect the 1904 calendar. | Export .NET DateTime values to Excel while preserving correct serial numbers. | Build templates where date formatting must follow the 1904 system. | Ensure cross‑platform date consistency when sharing files with Mac users.
// AI Prompts: Write C# code with Aspose.Cells to enable the 1904 date system and insert a formatted date into cell A1. | Show how to convert a .NET DateTime to an Excel serial number using the 1904 flag via CellsHelper in Aspose.Cells. | Explain the steps to guarantee that formulas dependent on the 1904 date system evaluate correctly after saving the workbook.

using System;
using Aspose.Cells;

// Create a new Workbook, activate the 1904 date system via Workbook.Settings.Date1904, convert a .NET DateTime to an Excel serial number with CellsHelper.GetDoubleFromDateTime, write the value to a cell, apply a built‑in date format, and save the file—ensuring compatibility with legacy Mac Excel date calculations.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook wb = new Workbook();

        // Enable the 1904 date system for legacy formulas
        wb.Settings.Date1904 = true;

        // Example: write a date value that respects the 1904 system
        Worksheet sheet = wb.Worksheets[0];
        Cell cell = sheet.Cells["A1"];
        DateTime date = new DateTime(2000, 1, 1);
        // Convert DateTime to Excel serial number using the 1904 flag
        double serialDate = CellsHelper.GetDoubleFromDateTime(date, true);
        cell.PutValue(serialDate);

        // Apply a standard date format to the cell
        Style style = cell.GetStyle();
        style.Number = 14; // Built‑in date format
        cell.SetStyle(style);

        // Save the workbook (lifecycle rule: save)
        wb.Save("1904DateSystem.xlsx");
    }
}
