using System;
using Aspose.Cells;

namespace AsposeCellsBackupExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (source workbook)
            Workbook workbook = new Workbook();

            // Access the first worksheet (source sheet) and add sample data
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";
            sourceSheet.Cells["A1"].PutValue("Item");
            sourceSheet.Cells["B1"].PutValue("Quantity");
            sourceSheet.Cells["A2"].PutValue("Apple");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["A3"].PutValue("Orange");
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["C5"].PutValue(DateTime.Now); // include a date cell

            // Add a new worksheet that will hold the backup copy
            Worksheet backupSheet = workbook.Worksheets.Add("BackupCopy");

            // Enumerate all cells in the source sheet and copy their values to the backup sheet
            foreach (Cell srcCell in sourceSheet.Cells)
            {
                // Skip empty cells to avoid unnecessary writes
                if (srcCell.Type == CellValueType.IsNull) continue;

                // Use the same cell address in the backup sheet
                Cell destCell = backupSheet.Cells[srcCell.Name];
                // Preserve the original value (including numbers, strings, dates, booleans)
                destCell.PutValue(srcCell.Value);
            }

            // Save the workbook with the backup sheet
            workbook.Save("WorkbookWithBackup.xlsx", SaveFormat.Xlsx);
        }
    }
}