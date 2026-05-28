using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class GroupNumericPivotField
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data: a category column and a numeric amount column
        worksheet.Cells["A1"].Value = "Category";
        worksheet.Cells["B1"].Value = "Amount";

        worksheet.Cells["A2"].Value = "A";
        worksheet.Cells["B2"].Value = 5;
        worksheet.Cells["A3"].Value = "B";
        worksheet.Cells["B3"].Value = 15;
        worksheet.Cells["A4"].Value = "C";
        worksheet.Cells["B4"].Value = 25;
        worksheet.Cells["A5"].Value = "D";
        worksheet.Cells["B5"].Value = 45;
        worksheet.Cells["A6"].Value = "E";
        worksheet.Cells["B6"].Value = 65;
        worksheet.Cells["A7"].Value = "F";
        worksheet.Cells["B7"].Value = 85;

        // Create a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B7", "D3", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add the numeric field (Amount) to the row area – this field will be grouped
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Amount");

        // Add the category field to the data area (count of items per group)
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Category");

        // Retrieve the row field that represents the numeric values
        PivotField amountField = pivotTable.RowFields[0];

        // Group the numeric values from 0 to 100 with an interval of 20.
        // The last parameter 'true' creates a new grouped field while keeping the original.
        amountField.GroupBy(0, 100, 20, true);

        // Refresh the pivot table to apply the grouping and calculate the results
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the grouped pivot table
        workbook.Save("GroupedNumericPivot.xlsx");
    }
}