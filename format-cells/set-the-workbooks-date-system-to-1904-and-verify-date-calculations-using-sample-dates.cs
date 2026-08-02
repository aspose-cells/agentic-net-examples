// Title: Aspose.Cells C# – Enable 1904 Date System and Validate Serial‑Date Conversions
// Description: Creates a workbook, activates the 1904 date system, writes a DateTime to a cell with a custom format, extracts the Excel serial number, converts it back to DateTime, demonstrates handling of negative serial values (dates before 1904), and saves the file as DateSystem1904Demo.xlsx.
// Keywords: Aspose.Cells 1904 date system | C# Excel serial number conversion | GetDoubleFromDateTime Aspose | GetDateTimeFromDouble Aspose | negative Excel date serial | custom date format cell Aspose | Mac Excel 1904 compatibility | date system setting workbook
// Common Searches: how to set 1904 date system in Aspose.Cells .NET | convert DateTime to Excel serial number using 1904 system | retrieve DateTime from negative serial number Aspose.Cells | apply custom date‑time format in Aspose.Cells C# | verify 1904 date system in saved Excel file
// Developer Intent: Activate the 1904 date system for a workbook and ensure accurate round‑trip conversion between DateTime objects and Excel serial numbers, including support for dates earlier than 1904.
// Use Cases: Maintain compatibility with Mac Excel files that use the 1904 date system. | Calculate the serial number for a given .NET DateTime and confirm the reverse conversion yields the original value. | Display and manipulate dates preceding 1904 by using negative serial numbers.
// AI Prompts: Generate C# code that enables the 1904 date system in an Aspose.Cells workbook and formats a cell with a custom date‑time pattern. | Show how to convert a .NET DateTime to an Excel serial number and back using Aspose.Cells, handling negative serial values for pre‑1904 dates. | Explain how to confirm that the 1904 date system flag is persisted after saving the workbook and opening it in Excel.

using System;
using Aspose.Cells;

namespace AsposeCellsDateSystemDemo
{
    // Creates a workbook, activates the 1904 date system, writes a DateTime to a cell with a custom format, extracts the Excel serial number, converts it back to DateTime, demonstrates handling of negative serial values (dates before 1904), and saves the file as DateSystem1904Demo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Enable the 1904 date system for the workbook
            wb.Settings.Date1904 = true;

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample date to test (January 1, 2023)
            DateTime sampleDate = new DateTime(2023, 1, 1, 12, 30, 0);

            // Put the DateTime value into cell A1
            cells["A1"].PutValue(sampleDate);

            // Apply a date-time format to display the value correctly
            Style style = cells["A1"].GetStyle();
            style.Custom = "yyyy-mm-dd hh:mm:ss";
            cells["A1"].SetStyle(style);

            // Retrieve the underlying double (serial number) for the date
            double serialNumber = CellsHelper.GetDoubleFromDateTime(sampleDate, wb.Settings.Date1904);
            Console.WriteLine($"Serial number for {sampleDate:yyyy-MM-dd HH:mm:ss} (1904 system): {serialNumber}");

            // Convert the serial number back to DateTime using the 1904 system
            DateTime convertedDate = CellsHelper.GetDateTimeFromDouble(serialNumber, wb.Settings.Date1904);
            Console.WriteLine($"Converted back to DateTime: {convertedDate:yyyy-MM-dd HH:mm:ss}");

            // Demonstrate a negative date (date before 1904-01-01)
            double negativeSerial = -10.5; // 10.5 days before the base date
            cells["A2"].PutValue(negativeSerial);
            Style negStyle = cells["A2"].GetStyle();
            negStyle.Custom = "yyyy-mm-dd hh:mm:ss";
            cells["A2"].SetStyle(negStyle);
            DateTime negativeDate = CellsHelper.GetDateTimeFromDouble(negativeSerial, wb.Settings.Date1904);
            Console.WriteLine($"Negative serial {negativeSerial} converts to: {negativeDate:yyyy-MM-dd HH:mm:ss}");

            // Save the workbook to verify the settings in Excel
            wb.Save("DateSystem1904Demo.xlsx");
        }
    }
}
