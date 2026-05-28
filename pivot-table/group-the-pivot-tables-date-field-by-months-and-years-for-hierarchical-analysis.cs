using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class GroupPivotDateByMonthYear
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data: a Date column and a Sales column
        worksheet.Cells["A1"].PutValue("Date");
        worksheet.Cells["B1"].PutValue("Sales");

        worksheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue(new DateTime(2023, 2, 20));
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["A4"].PutValue(new DateTime(2023, 3, 5));
        worksheet.Cells["B4"].PutValue(200);
        worksheet.Cells["A5"].PutValue(new DateTime(2023, 4, 10));
        worksheet.Cells["B5"].PutValue(250);
        worksheet.Cells["A6"].PutValue(new DateTime(2023, 5, 25));
        worksheet.Cells["B6"].PutValue(300);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B6", "E3", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Place the Date field in the row area and the Sales field in the data area
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Date column (index 0)
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column (index 1)

        // Retrieve the pivot field that represents the Date column
        PivotField dateField = pivotTable.RowFields[0];

        // Define the grouping range (full year) and the desired group types (Months and Years)
        DateTime startDate = new DateTime(2023, 1, 1);
        DateTime endDate   = new DateTime(2023, 12, 31);
        PivotGroupByType[] groupTypes = new PivotGroupByType[]
        {
            PivotGroupByType.Months,
            PivotGroupByType.Years
        };

        // Group the date field by months and years, using an interval of 1.
        // The last parameter (false) indicates that the first group type (Months) is not added as a separate field.
        dateField.GroupBy(startDate, endDate, groupTypes, 1, false);

        // Refresh and calculate the pivot table to apply the grouping
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the grouped pivot table
        workbook.Save("PivotGroupedByMonthYear.xlsx");
    }
}