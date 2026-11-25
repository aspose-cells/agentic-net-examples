// Title: C# example: unfreeze panes and reapply FreezePanes with updated row index after inserting rows using Aspose.Cells
// Common Searches: Aspose.Cells C# adjust FreezePanes after inserting rows | how to unfreeze and then reapply FreezePanes with new row index in Aspose.Cells | reset freeze panes to zero and set new FreezePanes position in Excel with Aspose.Cells | dynamic FreezePanes positioning after row insertion using Aspose.Cells for .NET | C# code to update frozen rows after adding rows in an Aspose.Cells workbook
// Tags: unfreeze panes Aspose.Cells C# | reapply FreezePanes with new row index Aspose.Cells | adjust FreezePanes after row insertion Aspose.Cells | dynamic FreezePanes positioning Aspose.Cells | reset freeze settings Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

namespace FreezePanesDemo
{
    // The sample loads or creates an Excel workbook, applies an initial FreezePanes at cell C3, inserts rows above the frozen area, clears the existing freeze by setting FreezePanes to zero, then reapplies FreezePanes using the shifted row index to keep the frozen rows aligned, and finally saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            try
            {
                // ------------------------------------------------------------
                // 1. Load an existing workbook if it exists; otherwise create a new one.
                // ------------------------------------------------------------
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine("Input file not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // ------------------------------------------------------------
                // 2. Access the first worksheet.
                // ------------------------------------------------------------
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine($"Working on worksheet: \"{sheet.Name}\"");

                // ------------------------------------------------------------
                // 3. Initial freeze: freeze at cell C3 (row index 2, column index 2 in zero‑based terms)
                //    Freeze 3 rows and 3 columns.
                // ------------------------------------------------------------
                int initialFreezeRow = 2;      // zero‑based index for row 3
                int initialFreezeColumn = 2;   // zero‑based index for column C
                int frozenRows = 3;
                int frozenColumns = 3;

                sheet.FreezePanes(initialFreezeRow, initialFreezeColumn, frozenRows, frozenColumns);
                Console.WriteLine($"Initial freeze applied at row {initialFreezeRow + 1}, column {initialFreezeColumn + 1}.");

                // ------------------------------------------------------------
                // 4. Insert rows above the frozen area to simulate data changes.
                //    Insert 2 rows before the original frozen row (row index 2).
                // ------------------------------------------------------------
                int rowsToInsert = 2;
                int insertAtRow = initialFreezeRow; // insert at the same index as the original freeze row
                sheet.Cells.InsertRows(insertAtRow, rowsToInsert);
                Console.WriteLine($"Inserted {rowsToInsert} rows at index {insertAtRow + 1}.");

                // ------------------------------------------------------------
                // 5. Unfreeze panes (remove any existing freeze).
                //    Aspose.Cells does not provide UnfreezePanes in some versions,
                //    so we reset the freeze settings to zero.
                // ------------------------------------------------------------
                sheet.FreezePanes(0, 0, 0, 0);
                Console.WriteLine("Panes unfrozen.");

                // ------------------------------------------------------------
                // 6. Re‑apply FreezePanes with updated indices.
                //    The original freeze row has moved down by the number of inserted rows.
                // ------------------------------------------------------------
                int newFreezeRow = initialFreezeRow + rowsToInsert;
                int newFreezeColumn = initialFreezeColumn; // column unchanged
                sheet.FreezePanes(newFreezeRow, newFreezeColumn, frozenRows, frozenColumns);
                Console.WriteLine($"Re‑applied freeze at row {newFreezeRow + 1}, column {newFreezeColumn + 1}.");

                // ------------------------------------------------------------
                // 7. Save the workbook.
                // ------------------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Report any exception; do not hide failures.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
