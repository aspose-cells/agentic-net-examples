using System;
using Aspose.Cells;

namespace AsposeCellsFilterExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header and several dates
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 20));
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 10));
            sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 5));
            sheet.Cells["A6"].PutValue(new DateTime(2023, 5, 12));
            sheet.Cells["A7"].PutValue(new DateTime(2023, 6, 30));

            // Define the auto‑filter range (including header and data rows)
            sheet.AutoFilter.Range = "A1:A7";

            // Apply a dynamic filter to keep only dates that belong to the current quarter
            // This automatically hides rows whose dates are outside the current quarter
            sheet.AutoFilter.DynamicFilter(0, DynamicFilterType.ThisQuarter);

            // Refresh the filter to apply the changes
            sheet.AutoFilter.Refresh();

            // Save the workbook
            workbook.Save("FilteredByCurrentQuarter.xlsx");
        }
    }
}