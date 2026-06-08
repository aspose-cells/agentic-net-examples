using System;
using Aspose.Cells;

namespace AsposeCellsDateFormatExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Enable the 1904 date system (compatible with the requirement)
            workbook.Settings.Date1904 = true;

            // Access the first worksheet and a target cell
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Put a sample date value into the cell
            DateTime sampleDate = new DateTime(2023, 5, 15); // May 15, 2023
            // Convert the DateTime to Excel serial number using the 1904 system
            double excelDate = CellsHelper.GetDoubleFromDateTime(sampleDate, workbook.Settings.Date1904);
            cell.PutValue(excelDate);

            // Define a custom number format that shows full month name and day (e.g., "May 15")
            Style style = cell.GetStyle();
            style.Custom = "mmmm d";
            cell.SetStyle(style);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("DateFormatFullMonthDay_1904.xlsx");
        }
    }
}