// Title: How to insert rows at a specific offset and import a DataTable into an Aspose.Cells worksheet using C#
// Common Searches: C# Aspose.Cells insert rows at row index 4 and shift existing data down | Import a DataTable into an Excel worksheet starting at a specific row with Aspose.Cells .NET | Set firstRow offset when adding a DataTable to a worksheet using Aspose.Cells | How to use InsertRows method to create space before importing DataTable in Aspose.Cells | Aspose.Cells C# example for inserting rows and loading DataTable with column headers
// Tags: insert rows with firstRow offset Aspose.Cells | DataTable import worksheet Aspose.Cells C# | shift existing rows down Aspose.Cells .NET | use InsertRows method Aspose.Cells | add column headers from DataTable Aspose.Cells | Excel row offset handling Aspose.Cells

using Aspose.Cells;
using System;
using System.Data;

// The sample creates a workbook, adds initial data, defines a DataTable, sets a firstRow offset of three, inserts rows to shift existing content down, writes the DataTable's column names and rows starting at the offset, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Existing data before the offset (rows 0‑2)
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Row1Col1");
            sheet.Cells["B2"].PutValue("Row1Col2");
            sheet.Cells["A3"].PutValue("Row2Col1");
            sheet.Cells["B3"].PutValue("Row2Col2");

            // Prepare the DataTable to import
            DataTable dt = new DataTable();
            dt.Columns.Add("Col1");
            dt.Columns.Add("Col2");
            dt.Rows.Add("Data1", "Data2");
            dt.Rows.Add("Data3", "Data4");
            dt.Rows.Add("Data5", "Data6");

            // Set the firstRow offset to three (zero‑based index)
            int firstRowOffset = 3;

            // Insert rows so existing rows shift down accordingly
            sheet.Cells.InsertRows(firstRowOffset, dt.Rows.Count + 1); // +1 for column names

            // Manually import column names
            int currentRow = firstRowOffset;
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                sheet.Cells[currentRow, c].PutValue(dt.Columns[c].ColumnName);
            }

            // Import the DataTable rows
            currentRow++; // start after column names
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    sheet.Cells[currentRow + r, c].PutValue(dt.Rows[r][c]);
                }
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
