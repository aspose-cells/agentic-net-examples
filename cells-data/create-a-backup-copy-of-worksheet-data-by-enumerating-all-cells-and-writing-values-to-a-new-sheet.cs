// Title: Create a backup worksheet by copying all populated cells from a source sheet using Aspose.Cells for .NET
// AI Prompts: Loop through each non‑empty cell in a source worksheet and write its value to the identical address on a newly added backup worksheet with Aspose.Cells in C#. | Generate a duplicate sheet in an Excel workbook by enumerating cells, copying their values, and saving the file as XLSX using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# copy only filled cells to another worksheet | How to programmatically duplicate a worksheet's data into a backup sheet in .NET | Iterate over cells in Aspose.Cells and write values to a new sheet before saving
// Tags: Aspose.Cells copy cell values between worksheets C# | enumerate non‑empty cells Aspose.Cells .NET | create additional worksheet Aspose.Cells Xlsx | add new worksheet and copy data Aspose.Cells | save workbook with multiple sheets Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsBackupExample
{
    // The example creates a workbook, adds sample data to a source worksheet, inserts a backup worksheet, iterates over all non‑empty cells of the source sheet copying each value to the same address in the backup sheet, and saves the workbook as WorkbookWithBackup.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook (source workbook)
                // -------------------------------------------------
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

                // -------------------------------------------------
                // 2. Add a new worksheet that will hold the backup copy
                // -------------------------------------------------
                Worksheet backupSheet = workbook.Worksheets.Add("Backup");

                // -------------------------------------------------
                // 3. Enumerate all cells in the source sheet and copy values
                // -------------------------------------------------
                foreach (Cell cell in sourceSheet.Cells)
                {
                    // Skip empty cells to avoid unnecessary writes
                    if (cell.Type != CellValueType.IsNull)
                    {
                        // Use the same address in the backup sheet
                        backupSheet.Cells[cell.Name].PutValue(cell.Value);
                    }
                }

                // -------------------------------------------------
                // 4. Save the workbook containing both original and backup sheets
                // -------------------------------------------------
                workbook.Save("WorkbookWithBackup.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
