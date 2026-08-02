// Title: Aspose.Cells C# – Enable 1904 Date System & Convert .NET DateTime to Excel Serial
// Description: Step‑by‑step guide to turn on the 1904 date system in an Aspose.Cells workbook, transform a .NET DateTime into an Excel serial number with CellsHelper, apply a date format, and save the file.
// Keywords: Aspose.Cells 1904 date system | C# workbook date base 1904 | Convert DateTime to Excel serial Aspose | CellsHelper GetDoubleFromDateTime | Excel legacy Mac date format | Date1904 property | Excel serial number conversion .NET | date number format Aspose.Cells
// Common Searches: how to set 1904 date system in Aspose.Cells | convert .NET DateTime to Excel serial number using Aspose | C# enable 1904 date base for Excel workbooks | apply built‑in date format after using 1904 system Aspose.Cells | legacy Mac Excel date handling with Aspose.Cells
// Developer Intent: Activate the 1904 date system and write .NET dates as Excel serial values in a C# workbook.
// Use Cases: Produce files compatible with older Mac Excel versions that rely on the 1904 epoch. | Export .NET DateTime data while preserving legacy serial numbers for existing formulas. | Generate reports where all date arithmetic must follow the 1904 base to match external data sources.
// AI Prompts: Show code to open an existing workbook and switch its date system to 1904 with Aspose.Cells. | Generate C# that converts a list of DateTime objects to 1904‑based Excel serial numbers and writes them to column B. | Explain how date calculations differ when wb.Settings.Date1904 is true in Aspose.Cells.

using System;
using Aspose.Cells;

// Step‑by‑step guide to turn on the 1904 date system in an Aspose.Cells workbook, transform a .NET DateTime into an Excel serial number with CellsHelper, apply a date format, and save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook wb = new Workbook();

        // Enable the 1904 date system for the workbook
        wb.Settings.Date1904 = true;

        // Example: convert a .NET DateTime to Excel serial number using the 1904 system
        DateTime sampleDate = new DateTime(2020, 1, 1);
        double excelSerial = CellsHelper.GetDoubleFromDateTime(sampleDate, true);

        // Write the serial value to a cell
        Worksheet sheet = wb.Worksheets[0];
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(excelSerial);

        // Apply a date number format so the value displays as a date
        Style style = cell.GetStyle();
        style.Number = 14; // Built‑in date format
        cell.SetStyle(style);

        // Save the workbook to a file
        wb.Save("1904DateSystem.xlsx");
    }
}
