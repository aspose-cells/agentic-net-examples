using System;
using Aspose.Cells;

namespace AsposeCellsCopyVisibleRowsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet srcSheet = workbook.Worksheets[0];
            Cells srcCells = srcSheet.Cells;

            // Populate header
            srcCells["A1"].PutValue("ID");
            srcCells["B1"].PutValue("Category");
            srcCells["C1"].PutValue("Value");

            // Populate sample data (rows 2‑11)
            for (int i = 0; i < 10; i++)
            {
                int row = i + 1; // zero‑based index (row 1 = second row)
                srcCells[row, 0].PutValue(i + 1);                     // ID
                srcCells[row, 1].PutValue(i % 2 == 0 ? "Even" : "Odd"); // Category
                srcCells[row, 2].PutValue(i * 10);                   // Value
            }

            // Apply AutoFilter to the header row covering columns A‑C
            srcSheet.AutoFilter.Range = "A1:C1";

            // Filter to show only rows where Value > 30
            srcSheet.AutoFilter.Custom(2, FilterOperatorType.GreaterThan, 30);
            srcSheet.AutoFilter.Refresh();

            // Determine number of data rows (excluding header)
            int totalDataRows = srcCells.MaxDisplayRange.RowCount - 1; // exclude header row

            // Set up copy options to copy only visible cells
            PasteOptions pasteOptions = new PasteOptions();
            pasteOptions.OnlyVisibleCells = true;

            // Default copy options (no special behavior)
            CopyOptions copyOptions = new CopyOptions();

            // Destination start row (e.g., row 15, zero‑based index 14)
            int destStartRow = 14;

            // Copy rows 2‑11 (sourceRowIndex = 1) to destination, respecting visibility
            srcCells.CopyRows(srcCells, 1, destStartRow, totalDataRows, copyOptions, pasteOptions);

            // Verify copied data: only visible rows should contain values
            Console.WriteLine("Verification of copied rows (only visible rows should have data):");
            for (int i = 0; i < totalDataRows; i++)
            {
                int srcRowIdx = i + 1;               // source data row index
                int destRowIdx = destStartRow + i;   // corresponding destination row index

                bool srcRowHidden = srcCells.IsRowHidden(srcRowIdx);
                string srcValue = srcCells[srcRowIdx, 0].StringValue;
                string destValue = srcCells[destRowIdx, 0].StringValue;

                Console.WriteLine($"Source Row {srcRowIdx + 1} (Hidden={srcRowHidden}) -> " +
                                  $"Dest Row {destRowIdx + 1}, Value='{destValue}' " +
                                  $"(Expected {(srcRowHidden ? "empty" : srcValue)})");
            }

            // Save the workbook to verify the result manually if needed
            workbook.Save("CopyVisibleRowsResult.xlsx");
        }
    }
}