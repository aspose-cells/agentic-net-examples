using System;
using Aspose.Cells;

namespace AsposeCellsColumnCopyDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create source workbook and populate a column with various data types
            Workbook srcWorkbook = new Workbook();
            Worksheet srcSheet = srcWorkbook.Worksheets[0];
            Cells srcCells = srcSheet.Cells;

            // Sample data in column B (index 1)
            srcCells[0, 1].PutValue(123);          // Integer
            srcCells[1, 1].PutValue(45.67);        // Double
            srcCells[2, 1].PutValue(true);         // Boolean
            srcCells[3, 1].PutValue(DateTime.Now); // DateTime
            srcCells[4, 1].PutValue("Text value"); // String

            // Set column width (in characters) for the source column
            srcSheet.Cells.SetColumnWidth(1, 25); // 25 characters wide

            // Create destination workbook (empty)
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            Cells destCells = destSheet.Cells;

            // Copy the entire column (data + formats) from source to destination
            // Parameters: source cells, source column index, destination column index
            destCells.CopyColumn(srcCells, 1, 0); // Copy source column B to destination column A

            // Preserve the column width by copying the width value explicitly
            double srcColumnWidth = srcSheet.Cells.GetColumnWidth(1);
            destSheet.Cells.SetColumnWidth(0, srcColumnWidth);

            // Save the result workbook
            destWorkbook.Save("ColumnCopyWithWidth.xlsx", SaveFormat.Xlsx);
        }
    }
}