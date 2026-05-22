using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class AutoFilterBetweenValuesDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with a header and some numeric amounts
        worksheet.Cells["A1"].PutValue("Amount");
        double[] amounts = { 120.5, 75.0, 200.0, 50.0, 180.0, 90.0 };
        for (int i = 0; i < amounts.Length; i++)
        {
            // Data starts from row 2 (index 1)
            worksheet.Cells[i + 1, 0].PutValue(amounts[i]);
        }

        // Define the auto‑filter range covering the header and data in column A
        // Parameters: start row (0), start column (0), end column (0) – column A only
        worksheet.AutoFilter.SetRange(0, 0, 0);

        // Apply a custom filter to show rows where Amount is between 80 and 180 (inclusive)
        double lowerBound = 80.0;
        double upperBound = 180.0;
        worksheet.AutoFilter.Custom(
            fieldIndex: 0,                         // Column A (Amount)
            operatorType1: FilterOperatorType.GreaterOrEqual,
            criteria1: lowerBound,
            isAnd: true,                           // Combine with AND
            operatorType2: FilterOperatorType.LessOrEqual,
            criteria2: upperBound);

        // Refresh the filter to hide rows that do not meet the criteria
        worksheet.AutoFilter.Refresh();

        // Save the workbook
        workbook.Save("FilteredAmounts.xlsx", SaveFormat.Xlsx);
    }
}