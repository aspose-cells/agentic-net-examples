// Title: Convert an Excel ListObject to a regular range, preserve formatting for the first five rows, and save as ODS with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert ListObject to range and keep first five rows formatting | How to clear formats after row 5 in a converted table using Aspose.Cells | Save workbook as ODS after table-to-range conversion in C# Aspose.Cells | Retain header and first few rows styling when converting Excel table to range with Aspose.Cells | Aspose.Cells example to export converted table to OpenDocument Spreadsheet
// Tags: ListObject to range conversion Aspose.Cells | preserve first five rows formatting Aspose.Cells | clear formats beyond specific rows Aspose.Cells | save workbook as ODS Aspose.Cells | C# Excel table conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel file, checks for a ListObject on the first worksheet, converts that table to a normal range, clears formatting for rows beyond the first five while keeping it for the initial rows, and finally saves the workbook in ODS format using Aspose.Cells for .NET.
class TableToRangeConverter
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.ods";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one table (ListObject) on the sheet
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found on the first worksheet.");
                return;
            }

            // Get the first table
            ListObject table = sheet.ListObjects[0];

            // Store table dimensions before conversion
            int firstRow = table.DataRange.FirstRow;          // zero‑based index of first data row
            int firstColumn = table.DataRange.FirstColumn;    // zero‑based index of first data column
            int totalRows = table.DataRange.RowCount;         // total number of data rows
            int totalColumns = table.DataRange.ColumnCount;   // total number of data columns

            // Convert the table to a normal range while preserving its data
            table.ConvertToRange();

            // Create a range object representing the converted data area
            AsposeRange dataRange = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);

            // Retain formatting only for the first five rows; clear styles for the rest
            if (totalRows > 5)
            {
                int rowsToClear = totalRows - 5;
                AsposeRange rowsBeyondFive = sheet.Cells.CreateRange(firstRow + 5, firstColumn, rowsToClear, totalColumns);
                // Clear only formatting (styles) from the rows beyond the first five
                rowsBeyondFive.ClearFormats();
            }

            // Save the workbook as ODS format
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
