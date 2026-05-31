using System;
using System.Collections;
using System.Text;
using System.Security.Cryptography;
using Aspose.Cells;

class RowChecksumExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (2 columns)
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue($"Item{i}");
            cells[i, 1].PutValue(i * 10);
        }

        // Determine the column where the checksum will be stored
        int checksumColumn = cells.MaxDataColumn + 1;
        cells[0, checksumColumn].PutValue("Checksum"); // header for checksum column

        // Enumerate all rows that contain data
        RowCollection rows = cells.Rows;
        IEnumerator rowEnumerator = rows.GetEnumerator();
        while (rowEnumerator.MoveNext())
        {
            Row row = (Row)rowEnumerator.Current;

            // Concatenate all cell values in the current row
            StringBuilder concatenated = new StringBuilder();
            IEnumerator cellEnumerator = row.GetEnumerator();
            while (cellEnumerator.MoveNext())
            {
                Cell cell = (Cell)cellEnumerator.Current;
                if (cell != null && cell.Value != null)
                {
                    concatenated.Append(cell.Value.ToString());
                }
            }

            // Compute SHA256 hash of the concatenated string
            string hashString;
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(concatenated.ToString()));
                hashString = BitConverter.ToString(hashBytes).Replace("-", "");
            }

            // Write the checksum into the designated column for this row
            cells[row.Index, checksumColumn].PutValue(hashString);
        }

        // Save the workbook with the computed checksums
        workbook.Save("RowChecksums.xlsx");
    }
}