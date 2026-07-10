using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotGroupingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data (Amount column)
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = 5;
            sheet.Cells["A3"].Value = 12;
            sheet.Cells["A4"].Value = 27;
            sheet.Cells["A5"].Value = 33;
            sheet.Cells["A6"].Value = 48;
            sheet.Cells["A7"].Value = 59;
            sheet.Cells["A8"].Value = 71;
            sheet.Cells["A9"].Value = 84;
            sheet.Cells["A10"].Value = 95;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:A10", "C3", "AmountPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Place the numeric field into the row area of the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);

            // Access the pivot field that was just added
            PivotField amountField = pivotTable.RowFields[0];

            // Group the numeric values into custom ranges:
            // Start = 0, End = 100, Interval = 20, do not create a new field (group in place)
            amountField.GroupBy(0, 100, 20, false);

            // Refresh and calculate the pivot table to apply the grouping
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Retrieve the numeric range group settings to demonstrate the applied interval
            PivotNumbericRangeGroupSettings groupSettings = amountField.GroupSettings as PivotNumbericRangeGroupSettings;
            if (groupSettings != null)
            {
                Console.WriteLine("Grouping applied:");
                Console.WriteLine($"Start = {groupSettings.Start}");
                Console.WriteLine($"End   = {groupSettings.End}");
                Console.WriteLine($"Interval = {groupSettings.Interval}");
            }

            // Save the workbook with the grouped pivot table
            workbook.Save("GroupedNumericPivot.xlsx");
        }
    }
}