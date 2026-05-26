using System;
using Aspose.Cells;

namespace AsposeCellsDate1904Demo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Enable the 1904 date system for the workbook (required setting)
            workbook.Settings.Date1904 = true;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Example date to insert
            DateTime sampleDate = new DateTime(2020, 1, 1, 12, 30, 0);

            // Convert the DateTime to Excel serial number using the 1904 system
            double excelDate = CellsHelper.GetDoubleFromDateTime(sampleDate, true);
            cells["A1"].PutValue(excelDate);

            // Apply a date format to display the value as a date/time
            Style style = cells["A1"].GetStyle();
            style.Number = 22; // Custom date/time format (e.g., "m/d/yyyy h:mm")
            cells["A1"].SetStyle(style);

            // Add a formula that depends on the date value (e.g., add 5 days)
            cells["B1"].Formula = "=A1 + 5";

            // Calculate formulas (optional, ensures B1 shows the correct result)
            workbook.CalculateFormula();

            // Save the workbook (save rule)
            workbook.Save("Date1904Demo.xlsx");
        }
    }
}