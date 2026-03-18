using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsColumnEnumeratorDemo
{
    // Enumerator that iterates through each column of a worksheet
    // and returns a list containing the values of that column (including nulls for empty cells).
    public class ColumnDataEnumerator : IEnumerator<IList<object>>
    {
        private readonly Worksheet _worksheet;
        private readonly int _maxColumnIndex;
        private readonly int _maxRowIndex;
        private int _currentColumn = -1;
        private IList<object> _currentColumnData;

        public ColumnDataEnumerator(Worksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            // Determine the range that actually contains data.
            _maxColumnIndex = worksheet.Cells.MaxDataColumn;
            _maxRowIndex = worksheet.Cells.MaxDataRow;
        }

        // Current column data (list of cell values)
        public IList<object> Current => _currentColumnData;

        object IEnumerator.Current => Current;

        // Move to the next column and collect its data.
        public bool MoveNext()
        {
            _currentColumn++;
            if (_currentColumn > _maxColumnIndex)
                return false;

            var columnData = new List<object>();
            for (int row = 0; row <= _maxRowIndex; row++)
            {
                Cell cell = _worksheet.Cells[row, _currentColumn];
                // Cell may be null if never instantiated; treat as empty.
                columnData.Add(cell?.Value);
            }

            _currentColumnData = columnData;
            return true;
        }

        public void Reset()
        {
            _currentColumn = -1;
            _currentColumnData = null;
        }

        public void Dispose()
        {
            // No unmanaged resources to release.
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and obtain the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data across several columns.
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");

            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["C2"].PutValue(30);

            sheet.Cells["A3"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            // Note: C3 left empty to demonstrate null handling.

            // Use the custom column enumerator.
            var columnEnumerator = new ColumnDataEnumerator(sheet);
            int columnIndex = 0;
            while (columnEnumerator.MoveNext())
            {
                IList<object> columnValues = columnEnumerator.Current;
                Console.WriteLine($"Column {columnIndex + 1}:");
                for (int i = 0; i < columnValues.Count; i++)
                {
                    Console.WriteLine($"  Row {i + 1}: {(columnValues[i] ?? "null")}");
                }
                columnIndex++;
            }

            // Save the workbook (demonstrating the required save rule).
            workbook.Save("ColumnEnumeratorDemo.xlsx");
        }
    }
}