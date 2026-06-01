using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDateGroupingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Populate sample data: Date column (A) and Sales column (B)
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Sales");

            // Sample dates spanning several months
            DateTime[] dates = new DateTime[]
            {
                new DateTime(2023, 1, 15),
                new DateTime(2023, 2, 20),
                new DateTime(2023, 3, 10),
                new DateTime(2023, 4, 5),
                new DateTime(2023, 5, 25),
                new DateTime(2023, 6, 12)
            };

            double[] sales = new double[] { 1500, 2300, 3200, 4100, 5000, 2750 };

            for (int i = 0; i < dates.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(dates[i]); // Column A
                sheet.Cells[i + 2, 1].PutValue(sales[i]); // Column B
            }

            // -------------------------------------------------
            // Create a pivot table based on the data range A1:B7
            // -------------------------------------------------
            int pivotIndex = sheet.PivotTables.Add("A1:B7", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Date field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");

            // Add the Sales field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // -------------------------------------------------
            // Group the Date field by Months and Years
            // -------------------------------------------------
            // Retrieve the date pivot field (first row field)
            PivotField dateField = pivotTable.RowFields[0];

            // Define the grouping range (cover all dates in the source data)
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate   = new DateTime(2023, 12, 31);

            // Specify the grouping types: Months and Years
            PivotGroupByType[] groupTypes = new PivotGroupByType[]
            {
                PivotGroupByType.Months,
                PivotGroupByType.Years
            };

            // Perform grouping: interval = 1 (default), do not create a new field
            dateField.GroupBy(startDate, endDate, groupTypes, 1, false);

            // Refresh and calculate the pivot table to apply the grouping
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // Save the workbook with the grouped pivot table
            // -------------------------------------------------
            workbook.Save("SalesPivotGroupedByMonthYear.xlsx");
        }
    }
}