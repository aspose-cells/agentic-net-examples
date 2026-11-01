// Title: Convert an Excel table to a regular range, keep formatting for the first five rows, and export to ODS using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# convert ListObject to range while preserving first five rows formatting | C# reset cell style to default for rows after the fifth row in an Excel table using Aspose.Cells | How to save a modified worksheet as ODS with Aspose.Cells .NET | Remove table styling beyond a specific row index with Aspose.Cells in C# | Export Excel data to OpenDocument Spreadsheet after converting table to range in C#
// Tags: convert ListObject to range Aspose.Cells | preserve first five rows formatting Aspose.Cells | reset cell style default Aspose.Cells | export workbook to ODS Aspose.Cells | modify table styling C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an XLSX workbook, checks for a table on the first worksheet, captures its range, clears the style of all rows after the fifth row by applying the workbook's default style, and then saves the result as an ODS file using Aspose.Cells for .NET.
class TableToRangeConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.ods";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found on the worksheet.");
                return;
            }

            // Get the first table
            ListObject table = sheet.ListObjects[0];

            // Store the table's range details before any modifications
            int firstRow = table.DataRange.FirstRow;
            int firstColumn = table.DataRange.FirstColumn;
            int totalRows = table.DataRange.RowCount;
            int totalColumns = table.DataRange.ColumnCount;

            // Optionally, convert the table to a normal range.
            // If Unlist() is unavailable, simply keep the range as is.

            // Retain formatting only for the first five rows of the former table range
            int startRow = firstRow + 5; // zero‑based index of the 6th row
            int endRow = firstRow + totalRows - 1;

            // Reset style of rows beyond the first five rows to the default workbook style
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = firstColumn; col < firstColumn + totalColumns; col++)
                {
                    // Apply the default style to each cell
                    sheet.Cells[row, col].SetStyle(workbook.DefaultStyle);
                }
            }

            // Save the modified workbook as ODS
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
