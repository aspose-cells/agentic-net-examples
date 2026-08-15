// Title: Apply a "Month Day" custom format with the 1904 date system using Aspose.Cells for .NET
// Description: Creates a new Workbook, switches to the 1904 date system, converts a .NET DateTime to the Excel serial value, writes it to cell A1, applies the custom number format "mmmm d" to display the full month name and day, and saves the file as DateFormat1904.xlsx.
// Keywords: Aspose.Cells | C# | 1904 date system | custom date format | full month name | Excel serial number | CellsHelper.GetDoubleFromDateTime | Excel Mac compatibility | number format mmmm d | date formatting in Aspose.Cells
// Common Searches: Aspose.Cells 1904 date system example | C# format cell as month name and day in Excel | Convert DateTime to Excel serial number 1904 Aspose | Apply custom number format mmmm d with Aspose.Cells | Enable 1904 date system in Aspose.Cells .NET
// Developer Intent: Generate an Excel workbook that uses the 1904 date system and displays dates as full month name plus day.
// Use Cases: Producing historical reports that require the 1904 date system (e.g., Mac‑compatible files). | Exporting data where a readable "July 15" style date is needed without losing serial precision. | Building templates that combine custom date formatting with correct serial values for legacy Excel versions.
// AI Prompts: Show C# code to enable the 1904 date system in Aspose.Cells and format a cell with "mmmm d". | How do I convert a .NET DateTime to an Excel serial number for the 1904 system using Aspose.Cells? | Explain the steps to apply a custom date format that displays the full month name and day in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// Creates a new Workbook, switches to the 1904 date system, converts a .NET DateTime to the Excel serial value, writes it to cell A1, applies the custom number format "mmmm d" to display the full month name and day, and saves the file as DateFormat1904.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Enable the 1904 date system
        wb.Settings.Date1904 = true;

        // Access the first worksheet and a target cell
        Worksheet sheet = wb.Worksheets[0];
        Cell cell = sheet.Cells["A1"];

        // Define the date to display (e.g., July 15, 2023)
        DateTime date = new DateTime(2023, 7, 15);

        // Convert the DateTime to Excel's serial number using the 1904 system
        double serial = CellsHelper.GetDoubleFromDateTime(date, true);
        cell.PutValue(serial);

        // Apply a custom format that shows the full month name and day
        Style style = cell.GetStyle();
        style.Custom = "mmmm d";
        cell.SetStyle(style);

        // Save the workbook
        wb.Save("DateFormat1904.xlsx");
    }
}
