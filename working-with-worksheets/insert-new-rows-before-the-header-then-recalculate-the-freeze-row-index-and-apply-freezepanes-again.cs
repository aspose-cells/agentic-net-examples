using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class InsertRowsAndReapplyFreezePanes
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Add sample data: a header row at the top and some data rows
                cells["A1"].PutValue("Header");
                cells["A2"].PutValue("Data 1");
                cells["A3"].PutValue("Data 2");
                cells["A4"].PutValue("Data 3");

                // Freeze panes initially at row 3 (zero‑based index) and column 0
                // This freezes the first three rows (including the header)
                worksheet.FreezePanes(3, 0, 3, 0);

                // Store the original freeze pane information
                int originalRow, originalColumn, originalFreezedRows, originalFreezedColumns;
                worksheet.GetFreezedPanes(out originalRow, out originalColumn, out originalFreezedRows, out originalFreezedColumns);

                // Number of rows to insert before the header
                int rowsToInsert = 2;

                // Insert rows at index 0 (before the header)
                cells.InsertRows(0, rowsToInsert);

                // Recalculate freeze pane positions after the insertion
                int newRow = originalRow + rowsToInsert;
                int newColumn = originalColumn; // column does not change
                int newFreezedRows = originalFreezedRows + rowsToInsert;
                int newFreezedColumns = originalFreezedColumns; // column count unchanged

                // Unfreeze any existing panes (optional but ensures clean state)
                worksheet.UnFreezePanes();

                // Apply the updated freeze panes
                worksheet.FreezePanes(newRow, newColumn, newFreezedRows, newFreezedColumns);

                // Save the workbook
                string outputPath = "InsertRowsAndReapplyFreezePanes.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            InsertRowsAndReapplyFreezePanes.Run();
        }
    }
}