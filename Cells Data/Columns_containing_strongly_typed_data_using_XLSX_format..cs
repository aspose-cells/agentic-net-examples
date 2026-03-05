using System;
using Aspose.Cells;

namespace StronglyTypedColumnsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add headers for each column
            sheet.Cells["A1"].PutValue("IntegerColumn");
            sheet.Cells["B1"].PutValue("DoubleColumn");
            sheet.Cells["C1"].PutValue("DateTimeColumn");
            sheet.Cells["D1"].PutValue("BooleanColumn");

            // Populate column A with integers (strongly typed)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(i * 10); // Row index i+1, column index 0 (A)
            }

            // Populate column B with doubles
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 1].PutValue(i * 0.5 + 1.25); // Column B
            }

            // Populate column C with DateTime values
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 2].PutValue(startDate.AddDays(i)); // Column C
            }

            // Populate column D with boolean values
            for (int i = 0; i < 10; i++)
            {
                bool flag = (i % 2 == 0);
                sheet.Cells[i + 1, 3].PutValue(flag); // Column D
            }

            // Optionally, auto-fit columns for better visibility
            sheet.AutoFitColumns();

            // Save the workbook in XLSX format (lifecycle: save)
            workbook.Save("StronglyTypedColumns.xlsx", SaveFormat.Xlsx);
        }
    }
}