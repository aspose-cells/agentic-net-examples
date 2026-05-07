using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeGroupingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Populate sample data in the worksheet (rows 0-9, columns 0-4)
            // ------------------------------------------------------------
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ------------------------------------------------------------
            // Create a range covering cells A1:D5 using CreateRange(string, string)
            // ------------------------------------------------------------
            AsposeRange dataRange = cells.CreateRange("A1", "D5");
            dataRange.Name = "DataBlock";

            // Set a header for the range
            dataRange[0, 0].PutValue("Header1");
            dataRange[0, 1].PutValue("Header2");
            dataRange[0, 2].PutValue("Header3");
            dataRange[0, 3].PutValue("Header4");

            // ------------------------------------------------------------
            // Add the created range to the worksheet's Cells collection
            // (so it expands automatically when rows/columns are inserted)
            // ------------------------------------------------------------
            cells.AddRange(dataRange);

            // ------------------------------------------------------------
            // Group rows 2 to 4 (zero‑based indices 1‑3) and hide them
            // ------------------------------------------------------------
            cells.GroupRows(1, 3, true);

            // ------------------------------------------------------------
            // Group columns 3 to 5 (zero‑based indices 2‑4) without hiding
            // ------------------------------------------------------------
            cells.GroupColumns(2, 3, false);

            // ------------------------------------------------------------
            // Clear the contents and formatting of a sub‑range (B2:C3)
            // ------------------------------------------------------------
            CellArea clearArea = CellArea.CreateCellArea("B2", "C3");
            cells.ClearRange(clearArea);

            // ------------------------------------------------------------
            // Insert a new blank row above row 6 and shift cells down,
            // updating any references automatically
            // ------------------------------------------------------------
            CellArea insertArea = CellArea.CreateCellArea(5, 0, 5, cells.MaxColumn);
            cells.InsertRange(insertArea, 1, ShiftType.Down, true);

            // ------------------------------------------------------------
            // Save the workbook to an XLSX file (lifecycle rule: save)
            // ------------------------------------------------------------
            workbook.Save("GroupedRangeDemo.xlsx");
        }
    }
}