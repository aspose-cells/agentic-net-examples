using System;
using Aspose.Cells;

namespace AsposeCellsLightCellsCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.25);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Add comments to a few cells
            int commentIdxA1 = sheet.Comments.Add("A1");
            Comment commentA1 = sheet.Comments[commentIdxA1];
            commentA1.Note = "Header for items";
            commentA1.Author = "User1";

            int commentIdxB2 = sheet.Comments.Add(1, 1); // B2 (row 1, column 1)
            Comment commentB2 = sheet.Comments[commentIdxB2];
            commentB2.Note = "Price of Apple";
            commentB2.Author = "User2";

            // Create a LightCellsDataProvider that streams existing cell values.
            // It does not modify comments; comments are stored in the workbook and will be saved.
            var dataProvider = new SimpleLightCellsDataProvider(sheet);

            // Set up save options with the LightCellsDataProvider
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = dataProvider
            };

            // Save the workbook; comments are preserved in the output file.
            workbook.Save("WorkbookWithComments.xlsx", saveOptions);
        }
    }

    // Simple implementation of LightCellsDataProvider that iterates over all used rows/columns.
    // It writes the existing cell values without altering other workbook objects (e.g., comments).
    public class SimpleLightCellsDataProvider : LightCellsDataProvider
    {
        private readonly Worksheet _sheet;
        private int _currentRow = -1;
        private int _currentColumn = -1;
        private readonly int _maxRow;
        private readonly int _maxColumn;

        public SimpleLightCellsDataProvider(Worksheet sheet)
        {
            _sheet = sheet;
            // Determine the used range to know how many rows/columns to process.
            var maxRow = sheet.Cells.MaxDataRow;
            var maxCol = sheet.Cells.MaxDataColumn;
            _maxRow = maxRow >= 0 ? maxRow : 0;
            _maxColumn = maxCol >= 0 ? maxCol : 0;
        }

        // Process only the first sheet (index 0)
        public bool StartSheet(int sheetIndex)
        {
            return sheetIndex == 0;
        }

        // Return the next row index or -1 when done.
        public int NextRow()
        {
            if (_currentRow < _maxRow)
            {
                _currentRow++;
                _currentColumn = -1; // reset column for new row
                return _currentRow;
            }
            return -1;
        }

        // No special row handling needed.
        public void StartRow(Row row) { }

        // Return the next column index for the current row or -1 when done.
        public int NextCell()
        {
            if (_currentColumn < _maxColumn)
            {
                _currentColumn++;
                return _currentColumn;
            }
            return -1;
        }

        // Preserve the existing cell value; no modification required.
        public void StartCell(Cell cell)
        {
            // The cell already contains its value; we simply leave it unchanged.
            // This ensures that comments attached to the cell are retained.
        }

        // Indicate whether string values should be gathered into a global pool.
        public bool IsGatherString()
        {
            return true;
        }
    }
}