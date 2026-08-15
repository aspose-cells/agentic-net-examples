// Title: C# Sample: Enable 1904 Date System in Aspose.Cells Workbook and Verify Date‑Serial Conversions
// Description: Demonstrates how to switch a workbook to the 1904 date system using Aspose.Cells for .NET, write DateTime values, format them, convert dates to Excel serial numbers (and back) with CellsHelper, output verification results, and save the file. Ideal for Mac‑compatible Excel files and date‑calculation testing.
// Keywords: Aspose.Cells 1904 date system | C# Excel date serial number | CellsHelper GetDoubleFromDateTime | Excel 1904 epoch .NET | date to serial conversion Aspose | verify Excel date calculations | Mac Excel compatibility | Aspose.Cells workbook settings | date round‑trip test | C# Excel automation
// Common Searches: how to set 1904 date system in Aspose.Cells | convert DateTime to Excel serial number C# Aspose | Aspose.Cells round‑trip date conversion example | 1904 epoch Excel date handling .NET | Aspose.Cells date format built‑in style | verify 1904 date system after saving workbook
// Developer Intent: Show how to activate the 1904 date system in an Aspose.Cells workbook and confirm accurate date‑to‑serial and serial‑to‑date conversions.
// Use Cases: Create workbooks compatible with older Mac Excel versions that require the 1904 epoch. | Store and retrieve dates as serial numbers while preserving correct values under the 1904 system. | Run a quick round‑trip test to ensure date calculations remain consistent before publishing a spreadsheet.
// AI Prompts: Provide C# code to enable the 1904 date system in an Aspose.Cells workbook and apply a built‑in date format. | Show how to convert a .NET DateTime to an Excel serial number and back using CellsHelper with the 1904 flag. | Explain steps to verify that the 1904 date system is correctly applied after saving the workbook.

using System;
using Aspose.Cells;

// Demonstrates how to switch a workbook to the 1904 date system using Aspose.Cells for .NET, write DateTime values, format them, convert dates to Excel serial numbers (and back) with CellsHelper, output verification results, and save the file. Ideal for Mac‑compatible Excel files and date‑calculation testing.
class DateSystem1904Demo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Enable the 1904 date system
        wb.Settings.Date1904 = true;

        // Access the first worksheet and its cells
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Sample dates to test
        DateTime baseDate = new DateTime(1904, 1, 1);          // Excel's 1904 epoch
        DateTime sampleDate = new DateTime(2023, 12, 31);     // Arbitrary later date

        // Put the DateTime values directly into cells
        cells["A1"].PutValue(baseDate);
        cells["A2"].PutValue(sampleDate);

        // Apply a built‑in date format so the values display as dates
        Style dateStyle = wb.CreateStyle();
        dateStyle.Number = 14; // Built‑in date format (e.g., mm/dd/yyyy)
        cells["A1"].SetStyle(dateStyle);
        cells["A2"].SetStyle(dateStyle);

        // Convert the same dates to Excel serial numbers using the 1904 system
        double serialBase = CellsHelper.GetDoubleFromDateTime(baseDate, true);
        double serialSample = CellsHelper.GetDoubleFromDateTime(sampleDate, true);

        // Store the serial numbers in column B and format them as dates
        cells["B1"].PutValue(serialBase);
        cells["B2"].PutValue(serialSample);
        cells["B1"].SetStyle(dateStyle);
        cells["B2"].SetStyle(dateStyle);

        // Convert the serial numbers back to DateTime to verify correctness
        DateTime backBase = CellsHelper.GetDateTimeFromDouble(serialBase, true);
        DateTime backSample = CellsHelper.GetDateTimeFromDouble(serialSample, true);

        // Output verification results to the console
        Console.WriteLine($"Original {baseDate:yyyy-MM-dd} -> Serial {serialBase} -> Back {backBase:yyyy-MM-dd}");
        Console.WriteLine($"Original {sampleDate:yyyy-MM-dd} -> Serial {serialSample} -> Back {backSample:yyyy-MM-dd}");

        // Save the workbook
        wb.Save("DateSystem1904Demo.xlsx");
    }
}
