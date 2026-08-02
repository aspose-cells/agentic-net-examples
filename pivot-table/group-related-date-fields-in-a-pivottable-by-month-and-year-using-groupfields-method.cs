using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDateGrouping
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a Date column and a numeric Value column
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");

            // Sample dates spanning several months
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 10));
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 5));
            sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 20));
            sheet.Cells["A6"].PutValue(new DateTime(2023, 5, 25));

            // Corresponding numeric values
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["B5"].PutValue(300);
            sheet.Cells["B6"].PutValue(250);

            // Create a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Date field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");

            // Add the Value field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Retrieve the PivotField representing the Date column
            PivotField dateField = pivotTable.RowFields[0];

            // Define the start and end dates for grouping
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate   = new DateTime(2023, 12, 31);

            // Group the date field by Months and Years
            // Interval is set to 1 (default), and firstAsNewField = false (adds grouping to existing field)
            dateField.GroupBy(startDate, endDate,
                new PivotGroupByType[] { PivotGroupByType.Months, PivotGroupByType.Years },
                1, false);

            // Refresh the pivot table to apply grouping
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotDateGroupedByMonthYear.xlsx");
        }
    }
}