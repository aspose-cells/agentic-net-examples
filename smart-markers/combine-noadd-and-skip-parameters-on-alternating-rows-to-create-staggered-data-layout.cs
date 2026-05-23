using System;
using Aspose.Cells;

namespace AsposeCellsStaggeredImportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data to be imported – each inner array represents a row of values
            object[][] rowsData = new object[][]
            {
                new object[] { "Item1", "Item2", "Item3" },
                new object[] { "Item4", "Item5", "Item6" },
                new object[] { "Item7", "Item8", "Item9" },
                new object[] { "Item10", "Item11", "Item12" }
            };

            // Import each row with alternating skip (noadd) behavior:
            //   * Even-indexed rows (0,2,…) are imported without column gaps (skip = 0)
            //   * Odd-indexed rows (1,3,…) are imported with a column gap of 1 (skip = 1)
            for (int rowIndex = 0; rowIndex < rowsData.Length; rowIndex++)
            {
                object[] currentRow = rowsData[rowIndex];
                int skip = (rowIndex % 2 == 0) ? 0 : 1; // alternate between 0 and 1

                // ImportObjectArray parameters:
                //   objArray   – data for the current row
                //   firstRow   – target row index in the worksheet
                //   firstColumn– start column (0 = column A)
                //   isVertical – false = horizontal import (across columns)
                //   skip       – number of columns to skip between each entry
                worksheet.Cells.ImportObjectArray(currentRow, rowIndex, 0, false, skip);
            }

            // Save the workbook to demonstrate the staggered layout
            workbook.Save("StaggeredDataLayout.xlsx");
        }
    }
}