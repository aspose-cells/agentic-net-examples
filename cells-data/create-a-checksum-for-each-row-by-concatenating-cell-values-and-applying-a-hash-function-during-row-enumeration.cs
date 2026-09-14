// Title: Create a checksum column for each row in an Excel sheet using Aspose.Cells for .NET by concatenating cell values and applying GetHashCode
// AI Prompts: Generate C# code that loops through all rows with Aspose.Cells, builds a pipe‑delimited string of each row’s cell values, calculates an integer hash with GetHashCode, and writes the result to the next empty column. | Show how to add a reusable method that accepts a Row object and returns a checksum based on concatenated cell contents using Aspose.Cells. | Demonstrate saving the workbook after inserting the computed checksum column into an existing Excel file with Aspose.Cells.
// Common Searches: aspnet compute checksum for each Excel row using Aspose.Cells | how to add a hash column to a worksheet with Aspose.Cells C# | concatenate row values and generate GetHashCode in Aspose.Cells | enumerate rows and write calculated checksum to new column in Excel via .NET | example of per‑row checksum generation with Aspose.Cells library
// Tags: row checksum Aspose.Cells | concatenate cell values hash .NET | add checksum column Excel | enumerate rows Aspose.Cells API | compute row hash C#

using System;
using System.Collections;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsChecksumDemo
{
    // The program creates a workbook, fills it with sample data, iterates each row using Aspose.Cells, concatenates the cell values with a delimiter, computes an integer checksum via GetHashCode, writes the checksum into the next empty column of the row, and saves the file as ChecksumDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (3 rows, 3 columns)
            cells["A1"].PutValue("John");
            cells["B1"].PutValue(28);
            cells["C1"].PutValue("Engineer");

            cells["A2"].PutValue("Alice");
            cells["B2"].PutValue(34);
            cells["C2"].PutValue("Manager");

            cells["A3"].PutValue("Bob");
            cells["B3"].PutValue(22);
            cells["C3"].PutValue("Analyst");

            // Enumerate through each row in the worksheet
            IEnumerator rowEnum = sheet.Cells.Rows.GetEnumerator();
            while (rowEnum.MoveNext())
            {
                Row row = (Row)rowEnum.Current;

                // Build a concatenated string of all cell values in the current row
                StringBuilder sb = new StringBuilder();
                IEnumerator cellEnum = row.GetEnumerator();
                while (cellEnum.MoveNext())
                {
                    Cell cell = (Cell)cellEnum.Current;
                    if (cell != null && cell.Value != null)
                    {
                        sb.Append(cell.Value.ToString());
                        sb.Append("|"); // delimiter to avoid accidental merging
                    }
                }

                // Compute a simple hash code from the concatenated string
                int checksum = sb.ToString().GetHashCode();

                // Determine the column index for the checksum (after the last existing cell)
                int checksumColumn = row.LastCell != null ? row.LastCell.Column + 1 : 0;

                // Write the checksum into the determined cell
                row[checksumColumn].PutValue(checksum);
            }

            // Save the workbook with the checksum column added
            workbook.Save("ChecksumDemo.xlsx");
        }
    }
}
