using System;
using Aspose.Cells;

namespace AsposeCellsMonthSheetsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Define month names
            string[] months = new string[]
            {
                "January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December"
            };

            // Loop through each month, add a worksheet, and populate sample data
            foreach (string month in months)
            {
                // Add a new worksheet with the month name (WorksheetCollection.Add(string) rule)
                Worksheet sheet = workbook.Worksheets.Add(month);

                // Add header row
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Description");
                sheet.Cells["C1"].PutValue("Amount");

                // Add a few sample rows
                sheet.Cells["A2"].PutValue($"{month} 1, 2023");
                sheet.Cells["B2"].PutValue("Sample expense");
                sheet.Cells["C2"].PutValue(100);

                sheet.Cells["A3"].PutValue($"{month} 15, 2023");
                sheet.Cells["B3"].PutValue("Sample income");
                sheet.Cells["C3"].PutValue(250);
            }

            // Save the workbook to disk (Save(string) rule)
            workbook.Save("MonthlySheets.xlsx");

            Console.WriteLine("Workbook with monthly worksheets created successfully.");
        }
    }
}