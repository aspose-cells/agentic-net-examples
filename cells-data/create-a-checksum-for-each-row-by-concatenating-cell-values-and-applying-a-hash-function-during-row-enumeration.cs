using System;
using System.Collections;
using System.Text;
using Aspose.Cells;

class RowChecksumDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Age");
        cells["A2"].PutValue("John");
        cells["B2"].PutValue(30);
        cells["A3"].PutValue("Alice");
        cells["B3"].PutValue(25);

        // Determine the column where the checksum will be stored (after the last used column)
        int checksumColumn = cells.MaxDataColumn + 1;
        cells[0, checksumColumn].PutValue("Checksum");

        // Enumerate all rows that contain data
        IEnumerator rowEnumerator = worksheet.Cells.Rows.GetEnumerator();
        while (rowEnumerator.MoveNext())
        {
            Row row = (Row)rowEnumerator.Current;

            // Concatenate the string representation of each cell's value in the current row
            StringBuilder concatenatedValues = new StringBuilder();
            IEnumerator cellEnumerator = row.GetEnumerator();
            while (cellEnumerator.MoveNext())
            {
                Cell cell = (Cell)cellEnumerator.Current;
                if (cell != null && cell.Value != null)
                {
                    concatenatedValues.Append(cell.Value.ToString());
                }
            }

            // Compute a simple hash code for the concatenated string
            int checksum = concatenatedValues.ToString().GetHashCode();

            // Write the checksum into the designated column of the current row
            cells[row.Index, checksumColumn].PutValue(checksum);
        }

        // Save the workbook with the checksum column added
        workbook.Save("RowChecksumDemo.xlsx");
    }
}