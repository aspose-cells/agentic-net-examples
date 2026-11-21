// Title: Insert a column at position five, copy the original column’s formatting, and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// Common Searches: how to insert a column at a specific index and preserve its formatting with Aspose.Cells in C# | Aspose.Cells C# copy column formatting after inserting a new column | export an Excel worksheet to PDF after adding a column using Aspose.Cells .NET | determine the last used row in a worksheet with Aspose.Cells C# | copy cell style from one column to another in Aspose.Cells example
// Tags: insert column copy formatting Aspose.Cells C# | duplicate column formatting after insertion Aspose.Cells | export worksheet to PDF Aspose.Cells | retrieve max data row Aspose.Cells | apply source cell style to destination Aspose.Cells

using System;
using Aspose.Cells;

// The code loads 'input.xlsx', inserts a new column at zero‑based index 4, copies the style from the original column (now at index 5) to the inserted column for all rows up to the maximum data row, and saves the workbook as 'output.pdf' in PDF format.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a new column at position five (zero‑based index 4)
        sheet.Cells.InsertColumn(4);

        // Determine the range of rows that contain data
        int maxRow = sheet.Cells.MaxDataRow;

        // Copy formatting from the original column (now at index 5) to the new column (index 4)
        for (int row = 0; row <= maxRow; row++)
        {
            // Source cell (original column before insertion)
            Cell sourceCell = sheet.Cells[row, 5];
            // Destination cell (newly inserted column)
            Cell destCell = sheet.Cells[row, 4];

            // Copy the style (formatting) only
            destCell.SetStyle(sourceCell.GetStyle());
        }

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
