using System;
using Aspose.Cells;

namespace AsposeCellsTop5FilterDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (header + 10 numeric rows)
            sheet.Cells["A1"].PutValue("Score");
            for (int i = 2; i <= 11; i++)
            {
                // Example values; you can replace with your own data
                sheet.Cells[i - 1, 0].PutValue(100 - (i - 2) * 7);
            }

            // Define the autofilter range (including header)
            sheet.AutoFilter.Range = "A1:A11";

            // Apply a Top 5 filter on the first column (field index 0)
            // Parameters: fieldIndex, isTop, isPercent, itemCount
            sheet.AutoFilter.FilterTop10(fieldIndex: 0, isTop: true, isPercent: false, itemCount: 5);

            // Refresh the filter to hide rows that do not meet the criteria
            sheet.AutoFilter.Refresh();

            // Save the workbook
            workbook.Save("Top5FilterDemo.xlsx");
        }
    }
}