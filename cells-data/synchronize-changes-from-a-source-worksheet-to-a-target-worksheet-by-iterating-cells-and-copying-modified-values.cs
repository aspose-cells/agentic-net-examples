using System;
using Aspose.Cells;

namespace AsposeCellsSyncDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (the workbook that contains the latest changes)
            Workbook sourceWorkbook = new Workbook("Source.xlsx");

            // Load the target workbook (the workbook that needs to be synchronized)
            Workbook targetWorkbook = new Workbook("Target.xlsx");

            // Assume we are synchronizing the first worksheet in each workbook.
            // Adjust the index or name as needed.
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Worksheet targetSheet = targetWorkbook.Worksheets[0];

            // Determine the used range of the source sheet.
            // This ensures we only iterate over cells that actually contain data.
            int maxRow = sourceSheet.Cells.MaxDisplayRange.RowCount;
            int maxColumn = sourceSheet.Cells.MaxDisplayRange.ColumnCount;

            // Iterate through each cell in the used range.
            for (int row = 0; row < maxRow; row++)
            {
                for (int col = 0; col < maxColumn; col++)
                {
                    // Retrieve the source and target cells.
                    var sourceCell = sourceSheet.Cells[row, col];
                    var targetCell = targetSheet.Cells[row, col];

                    // Compare the values. Use object.Equals to handle different data types safely.
                    // If the source cell is empty (null) and the target cell has a value, we also treat it as a change.
                    bool sourceHasValue = sourceCell.Value != null;
                    bool targetHasValue = targetCell.Value != null;

                    bool valuesDiffer = false;

                    if (sourceHasValue && targetHasValue)
                    {
                        valuesDiffer = !object.Equals(sourceCell.Value, targetCell.Value);
                    }
                    else if (sourceHasValue != targetHasValue)
                    {
                        // One cell is empty while the other is not.
                        valuesDiffer = true;
                    }

                    // If the values differ, copy the source value to the target cell.
                    if (valuesDiffer)
                    {
                        // Preserve the original cell type (string, number, date, etc.) by using PutValue.
                        targetCell.PutValue(sourceCell.Value);
                    }
                }
            }

            // Save the synchronized target workbook.
            targetWorkbook.Save("Target_Synchronized.xlsx");
        }
    }
}