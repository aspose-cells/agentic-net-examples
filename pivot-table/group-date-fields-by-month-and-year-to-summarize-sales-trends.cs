using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDateGrouping
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Populate sample data (Date and Sales) -----
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Sales");

            // Sample dates for the first half of 2023
            DateTime[] dates = new DateTime[]
            {
                new DateTime(2023, 1, 15),
                new DateTime(2023, 2, 20),
                new DateTime(2023, 3, 10),
                new DateTime(2023, 4, 5),
                new DateTime(2023, 5, 25),
                new DateTime(2023, 6, 12)
            };

            // Corresponding sales figures
            double[] sales = new double[] { 1500, 2300, 3200, 4100, 5000, 6100 };

            // Fill the worksheet with the data
            for (int i = 0; i < dates.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(dates[i]); // Column A (Date)
                sheet.Cells[i + 1, 1].PutValue(sales[i]); // Column B (Sales)
            }

            // ----- Create a pivot table based on the data range -----
            // Data range: A1:B7 (including header row)
            int pivotIndex = sheet.PivotTables.Add("A1:B7", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Date field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");

            // Add the Sales field to the data area (default aggregation is Sum)
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // ----- Group the Date field by Months and Years -----
            // Retrieve the date field from the row fields collection
            PivotField dateField = pivotTable.RowFields[0];

            // Define the grouping range (full year 2023) and the desired group types
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate   = new DateTime(2023, 12, 31);
            PivotGroupByType[] groupTypes = new PivotGroupByType[]
            {
                PivotGroupByType.Months,
                PivotGroupByType.Years
            };

            // Group by the specified range, using an interval of 1 and do not create a new field
            dateField.GroupBy(startDate, endDate, groupTypes, 1, false);

            // Refresh and calculate the pivot table to apply the grouping
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ----- Save the workbook -----
            workbook.Save("SalesPivotGroupedByMonthYear.xlsx");
        }
    }
}