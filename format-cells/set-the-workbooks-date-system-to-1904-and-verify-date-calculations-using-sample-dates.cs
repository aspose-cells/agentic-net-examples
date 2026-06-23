using System;
using Aspose.Cells;

namespace AsposeCellsDate1904Demo
{
    class Program
    {
        static void Main()
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

            // Demonstrate a negative date (date before 1904-01-01) using the 1904 system
            double negativeSerial = -10.5; // 10.5 days before 1904-01-01
            cells["A2"].PutValue(negativeSerial);
            Style negStyle = cells["A2"].GetStyle();
            negStyle.Custom = "yyyy-mm-dd hh:mm:ss";
            cells["A2"].SetStyle(negStyle);

            // Convert the negative serial to DateTime
            DateTime negativeDate = CellsHelper.GetDateTimeFromDouble(negativeSerial, wb.Settings.Date1904);
            Console.WriteLine($"Negative serial {negativeSerial} converts to: {negativeDate:yyyy-MM-dd HH:mm:ss}");

            // Save the workbook to verify the displayed values
            wb.Save("Date1904Demo.xlsx");
        }
    }
}