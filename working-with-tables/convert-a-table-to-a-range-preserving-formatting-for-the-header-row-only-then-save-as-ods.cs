// Title: Convert an Excel table to a plain range, keep only the header row formatting, and save as ODS using Aspose.Cells for .NET
// AI Prompts: Remove the first ListObject from a worksheet, retain its header row style, reset all data rows to the default style, and export the workbook to ODS with Aspose.Cells. | Transform an Excel table into a regular cell range, copy the original header formatting to the header cells, clear formatting from data rows, and save the file as an OpenDocument Spreadsheet.
// Common Searches: Aspose.Cells .NET convert ListObject to range while keeping header formatting | How to delete an Excel table but keep the header style using Aspose.Cells | Save workbook as ODS after removing table definition in C# | Preserve header row formatting when converting Excel table to plain cells with Aspose.Cells | C# code to remove table and export to OpenDocument Spreadsheet
// Tags: convert ListObject to range Aspose.Cells | header style preservation after table removal | save workbook to ODS format C# | reset cell styles for data rows Aspose.Cells | programmatic table deletion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an Excel file, checks for a table, removes the first ListObject while storing the header row's style, reapplies that style to the header cells, resets all data rows to the default workbook style, and saves the result as an ODS file.
class TableToRangeConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.ods";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the table
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one table (ListObject) on the sheet
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No table found on the worksheet.");
                return;
            }

            // Get the first ListObject (table)
            ListObject table = sheet.ListObjects[0];

            // Retrieve the cell area occupied by the table (including header)
            int headerRow = table.StartRow;          // Header row index
            int firstCol = table.StartColumn;        // First column index
            int lastCol = table.EndColumn;           // Last column index
            int lastRow = table.EndRow;              // Last row index (includes header)

            // Remove the table definition but keep the cells (converts to a plain range)
            sheet.ListObjects.RemoveAt(0);

            // Preserve formatting only for the header row
            Style headerStyle = sheet.Cells[headerRow, firstCol].GetStyle();

            // Apply the stored header style to all header cells
            for (int col = firstCol; col <= lastCol; col++)
            {
                sheet.Cells[headerRow, col].SetStyle(headerStyle);
            }

            // Clear formatting for all data rows (rows after the header)
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                for (int col = firstCol; col <= lastCol; col++)
                {
                    sheet.Cells[row, col].SetStyle(workbook.DefaultStyle);
                }
            }

            // Save the modified workbook as ODS
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
