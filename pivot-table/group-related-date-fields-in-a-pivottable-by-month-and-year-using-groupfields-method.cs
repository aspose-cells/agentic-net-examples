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

            // Populate sample data with a Date column and a numeric value column
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Sales");

            // Sample dates spanning several months
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 10));
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 5));
            sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 20));
            sheet.Cells["A6"].PutValue(new DateTime(2023, 5, 25));

            // Corresponding sales values
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["B4"].PutValue(1800);
            sheet.Cells["B5"].PutValue(2100);
            sheet.Cells["B6"].PutValue(2400);

            // Create a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Date field to the row area and Sales to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the PivotField representing the Date column
            PivotField dateField = pivotTable.RowFields[0];

            // Define the grouping range (start and end dates)
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate   = new DateTime(2023, 12, 31);

            // Group by Months and Years with an interval of 1 (default)
            dateField.GroupBy(startDate, endDate,
                new PivotGroupByType[] { PivotGroupByType.Months, PivotGroupByType.Years },
                1, false);

            // Refresh and calculate the pivot table to apply grouping
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotDateGroupedByMonthYear.xlsx");
        }
    }
}