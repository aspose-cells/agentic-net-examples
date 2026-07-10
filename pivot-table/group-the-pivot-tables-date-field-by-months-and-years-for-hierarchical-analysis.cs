using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class PivotDateGroupExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Populate sample data -----
        // Header row
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Sales");

        // Sample dates and sales values
        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 20));
        sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 10));
        sheet.Cells["B4"].PutValue(1800);
        sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 5));
        sheet.Cells["B5"].PutValue(2100);
        sheet.Cells["A6"].PutValue(new DateTime(2023, 5, 25));
        sheet.Cells["B6"].PutValue(2400);
        sheet.Cells["A7"].PutValue(new DateTime(2023, 6, 30));
        sheet.Cells["B7"].PutValue(2700);
        sheet.Cells["A8"].PutValue(new DateTime(2023, 7, 15));
        sheet.Cells["B8"].PutValue(3000);
        sheet.Cells["A9"].PutValue(new DateTime(2023, 8, 20));
        sheet.Cells["B9"].PutValue(3300);
        sheet.Cells["A10"].PutValue(new DateTime(2023, 9, 10));
        sheet.Cells["B10"].PutValue(3600);
        sheet.Cells["A11"].PutValue(new DateTime(2023, 10, 5));
        sheet.Cells["B11"].PutValue(3900);
        sheet.Cells["A12"].PutValue(new DateTime(2023, 11, 25));
        sheet.Cells["B12"].PutValue(4200);
        sheet.Cells["A13"].PutValue(new DateTime(2023, 12, 31));
        sheet.Cells["B13"].PutValue(4500);

        // ----- Create a pivot table -----
        // Data range includes headers (A1:B13)
        int pivotIndex = sheet.PivotTables.Add("A1:B13", "D3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the Date field to the Row area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");

        // Add the Sales field to the Data area
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // ----- Group the Date field by Months and Years -----
        // Retrieve the date field (first row field)
        PivotField dateField = pivotTable.RowFields[0];

        // Define the grouping range (full year) and the desired group types
        DateTime startDate = new DateTime(2023, 1, 1);
        DateTime endDate   = new DateTime(2023, 12, 31);
        PivotGroupByType[] groups = new PivotGroupByType[]
        {
            PivotGroupByType.Months,
            PivotGroupByType.Years
        };

        // Perform grouping; interval set to 1 (default), first group becomes a new field = false
        dateField.GroupBy(startDate, endDate, groups, 1, false);

        // Refresh and calculate the pivot table to apply grouping
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // ----- Save the workbook -----
        workbook.Save("PivotDateGroupedByMonthsAndYears.xlsx");
    }
}