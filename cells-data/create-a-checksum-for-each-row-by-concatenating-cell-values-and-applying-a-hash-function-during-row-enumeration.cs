// Title: Aspose.Cells C# – Add a checksum column per row by hashing concatenated cell values
// Description: The sample builds a workbook, fills it with sample data, inserts a "Checksum" header, walks through every row using the Rows enumerator, joins the text of each non‑null cell, creates an integer hash with GetHashCode, places the result in the first free column of that row, and writes the file to disk.
// Keywords: Aspose.Cells | C# | row checksum | hash code | concatenate cells | enumerate rows | Excel data integrity | SHA‑256 alternative | checksum column | GetHashCode | worksheet validation | global
// Common Searches: how to generate a checksum for each Excel row using Aspose.Cells C# | store hash of row data in a new column with Aspose.Cells | skip header row while computing row hash in .NET | determine last used column in Aspose.Cells to add extra data | verify row integrity in an exported workbook
// Developer Intent: Create a per‑row hash from cell contents and write it to a new column in the same sheet.
// Use Cases: Validate that exported data has not been altered by comparing saved hashes with freshly computed ones. | Detect row‑level modifications in a worksheet for audit or synchronization processes. | Generate deterministic identifiers for rows to support deduplication or fast look‑ups.
// AI Prompts: Write C# code with Aspose.Cells that computes a SHA‑256 hash of all cell values in each row and saves it in a new column. | Show how to exclude the header row and ignore empty cells when building a row‑level checksum in Aspose.Cells. | Explain a method to read back the checksum column later and flag rows whose current hash differs from the stored value.

using System;
using System.Collections;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsChecksumDemo
{
    // The sample builds a workbook, fills it with sample data, inserts a "Checksum" header, walks through every row using the Rows enumerator, joins the text of each non‑null cell, creates an integer hash with GetHashCode, places the result in the first free column of that row, and writes the file to disk.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including a header row)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["C1"].PutValue("Country");

            cells["A2"].PutValue("John");
            cells["B2"].PutValue(30);
            cells["C2"].PutValue("USA");

            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(25);
            cells["C3"].PutValue("UK");

            // Add a header for the checksum column
            cells["D1"].PutValue("Checksum");

            // Iterate through each row using the RowCollection enumerator
            IEnumerator rowEnum = sheet.Cells.Rows.GetEnumerator();
            while (rowEnum.MoveNext())
            {
                Row row = (Row)rowEnum.Current;

                // Skip the header row (index 0) if you don't want to checksum it
                // Uncomment the following line to skip:
                // if (row.Index == 0) continue;

                // Concatenate string representations of all cells in the row
                StringBuilder sb = new StringBuilder();
                IEnumerator cellEnum = row.GetEnumerator();
                while (cellEnum.MoveNext())
                {
                    Cell cell = (Cell)cellEnum.Current;
                    if (cell != null && cell.Value != null)
                    {
                        sb.Append(cell.Value.ToString());
                    }
                }

                // Compute a hash code from the concatenated string
                int checksum = sb.ToString().GetHashCode();

                // Determine the column index for the checksum (after the last existing cell)
                int checksumCol = row.LastCell != null ? row.LastCell.Column + 1 : 0;

                // Write the checksum value into the cell
                cells[row.Index, checksumCol].PutValue(checksum);
            }

            // Save the workbook
            workbook.Save("ChecksumDemo.xlsx");
        }
    }
}
