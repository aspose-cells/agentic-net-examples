using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsGroupFieldDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Assume the worksheet already contains a pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Example 1: Group a numeric field by interval
            // Add the numeric field to the row area (replace "Amount" with your field name)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Amount");
            // Retrieve the added field (it will be the last row field)
            PivotField numericField = pivotTable.RowFields[pivotTable.RowFields.Count - 1];
            // Group the numeric field by an interval of 10 (auto range, new field = false)
            numericField.GroupBy(10.0, false);

            // Example 2: Group a date field by months and years
            // Add the date field to the row area (replace "OrderDate" with your field name)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "OrderDate");
            // Retrieve the added date field (it will be the last row field now)
            PivotField dateField = pivotTable.RowFields[pivotTable.RowFields.Count - 1];
            // Define start and end dates for grouping
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate   = new DateTime(2023, 12, 31);
            // Group by months and years with an interval of 1
            dateField.GroupBy(startDate, endDate,
                new PivotGroupByType[] { PivotGroupByType.Months, PivotGroupByType.Years },
                1, false);

            // Refresh and calculate the pivot table to apply grouping
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");

            Console.WriteLine("Grouping applied and workbook saved successfully.");
        }
    }
}