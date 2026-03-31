using System;
using Aspose.Cells;

namespace AsposeCellsGroupAndRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range that contains the data to be subtotaled.
            // Adjust the start/end addresses according to your data layout.
            CellArea dataRange = CellArea.CreateCellArea("A1", "C10");

            // Specify the column index (zero‑based) to group by.
            // For example, column A (index 0) will be used as the grouping field.
            int groupByColumnIndex = 0;

            // Choose the aggregation function for the subtotal (e.g., Sum).
            ConsolidationFunction subtotalFunction = ConsolidationFunction.Sum;

            // Specify which columns (by zero‑based index) should receive the subtotal.
            // Here we subtotal the values in column C (index 2).
            int[] totalColumns = new int[] { 2 };

            // Apply the Subtotal operation using the defined range and group parameters.
            worksheet.Cells.Subtotal(
                dataRange,               // RANGE parameter
                groupByColumnIndex,      // GROUP parameter
                subtotalFunction,
                totalColumns
            );

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}