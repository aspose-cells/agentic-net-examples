using System;
using Aspose.Cells;

namespace AsposeCellsCopyRowsWithMergedCells
{
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook and add merged cells ----------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Merge cells A1:B2 (rows 0-1, columns 0-1) and put a value
            sourceCells.Merge(0, 0, 2, 2);
            sourceCells["A1"].PutValue("Merged Header");

            // Add some data in the first three rows
            sourceCells["A3"].PutValue("Row3-ColA");
            sourceCells["B3"].PutValue("Row3-ColB");
            sourceCells["A4"].PutValue("Row4-ColA");
            sourceCells["B4"].PutValue("Row4-ColB");

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            Cells destCells = destSheet.Cells;

            // ---------- Copy rows (including merged cells) ----------
            // Copy rows 0-3 (4 rows) from source to destination starting at row index 5
            int sourceStartRow = 0;
            int destinationStartRow = 5;
            int rowsToCopy = 4; // rows 0,1,2,3

            destCells.CopyRows(sourceCells, sourceStartRow, destinationStartRow, rowsToCopy);

            // ---------- Verify merged regions are preserved ----------
            // Get merged areas in the destination sheet
            CellArea[] mergedAreas = destSheet.Cells.GetMergedAreas();

            Console.WriteLine($"Number of merged areas in destination: {mergedAreas.Length}");
            foreach (CellArea area in mergedAreas)
            {
                // Display the address of each merged area
                string startAddress = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                string endAddress = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn);
                Console.WriteLine($"Merged area: {startAddress}:{endAddress}");
            }

            // Additional check: verify that the top-left cell of the merged area is indeed merged
            // The original merged area was A1:B2; after copying to start at row 5, it should be A6:B7
            Cell mergedCell = destSheet.Cells["A6"];
            Console.WriteLine($"Cell A6 IsMerged: {mergedCell.IsMerged}");

            // ---------- Save the destination workbook ----------
            destWorkbook.Save("CopiedRowsWithMergedCells.xlsx");
        }
    }
}