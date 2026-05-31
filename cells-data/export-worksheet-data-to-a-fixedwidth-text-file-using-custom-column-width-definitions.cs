using System;
using System.IO;
using Aspose.Cells;

class ExportFixedWidth
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["C1"].PutValue("Salary");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("John Doe");
        worksheet.Cells["C2"].PutValue(5230.75);
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Jane Smith");
        worksheet.Cells["C3"].PutValue(6100.00);

        // Define custom column widths (in character units)
        worksheet.Cells.Columns[0].Width = 5;   // ID column
        worksheet.Cells.Columns[1].Width = 20;  // Name column
        worksheet.Cells.Columns[2].Width = 10;  // Salary column

        // Determine the range of data to export
        int firstRow = 0;
        int lastRow = worksheet.Cells.MaxDataRow;
        int firstCol = 0;
        int lastCol = worksheet.Cells.MaxDataColumn;

        // Export the data to a fixed‑width text file
        using (StreamWriter writer = new StreamWriter("ExportFixedWidth.txt"))
        {
            for (int row = firstRow; row <= lastRow; row++)
            {
                string line = string.Empty;

                for (int col = firstCol; col <= lastCol; col++)
                {
                    // Get the cell's displayed text
                    string cellText = worksheet.Cells[row, col].StringValue;

                    // Get the column width in characters and round up
                    int colWidth = (int)Math.Ceiling(worksheet.Cells.GetColumnWidth(col));

                    // Truncate or pad the text to fit the column width
                    if (cellText.Length > colWidth)
                        cellText = cellText.Substring(0, colWidth);
                    else
                        cellText = cellText.PadRight(colWidth);

                    line += cellText;
                }

                writer.WriteLine(line);
            }
        }

        // Save the workbook (optional, for verification)
        workbook.Save("Reference.xlsx");
    }
}