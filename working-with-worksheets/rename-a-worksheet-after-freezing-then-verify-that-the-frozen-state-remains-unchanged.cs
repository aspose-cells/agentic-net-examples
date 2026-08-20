// Title: Rename an Excel worksheet after freezing panes and verify freeze state – Aspose.Cells C#
// Description: Demonstrates how to freeze panes at cell C3, capture the freeze configuration with GetFreezedPanes, rename the worksheet, re‑query the settings, compare before‑and‑after values, and save the workbook, proving that renaming does not alter frozen rows or columns.
// Keywords: Aspose.Cells freeze panes | rename worksheet C# | GetFreezedPanes example | preserve frozen rows after rename | Excel worksheet rename Aspose.Cells
// Common Searches: keep frozen panes when renaming a sheet Aspose.Cells | GetFreezedPanes after worksheet rename | C# freeze panes then rename worksheet | verify freeze state Aspose.Cells .NET
// Developer Intent: Rename a worksheet that has frozen panes and confirm the freeze configuration stays the same.
// Use Cases: Create a report sheet, freeze header rows/columns, rename for clarity, and ensure view settings persist. | Generate multiple worksheets programmatically, apply distinct freeze panes, rename each, and validate that all freeze settings remain intact. | Build an automated process where users rename worksheets on the fly while the application maintains frozen pane positions.
// AI Prompts: Write C# code with Aspose.Cells to freeze panes at a given cell, rename the worksheet, and check that the freeze state is unchanged. | Explain the GetFreezedPanes method and how to compare its output before and after a worksheet rename. | Suggest error‑handling strategies if the freeze configuration differs after renaming a sheet using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to freeze panes at cell C3, capture the freeze configuration with GetFreezedPanes, rename the worksheet, re‑query the settings, compare before‑and‑after values, and save the workbook, proving that renaming does not alter frozen rows or columns.
    public class RenameWorksheetAfterFreezeDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze panes at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns
            sheet.FreezePanes("C3", 2, 2);

            // Capture the freeze state before renaming
            bool isFrozenBefore = sheet.GetFreezedPanes(out int rowBefore, out int colBefore,
                                                       out int frozenRowsBefore, out int frozenColsBefore);
            Console.WriteLine($"Before rename - Frozen: {isFrozenBefore}, Row: {rowBefore}, Column: {colBefore}, " +
                              $"FrozenRows: {frozenRowsBefore}, FrozenColumns: {frozenColsBefore}");

            // Rename the worksheet
            string originalName = sheet.Name;
            string newName = "RenamedSheet";
            sheet.Name = newName;
            Console.WriteLine($"Worksheet renamed from '{originalName}' to '{sheet.Name}'");

            // Verify that the freeze state is unchanged after renaming
            bool isFrozenAfter = sheet.GetFreezedPanes(out int rowAfter, out int colAfter,
                                                      out int frozenRowsAfter, out int frozenColsAfter);
            Console.WriteLine($"After rename - Frozen: {isFrozenAfter}, Row: {rowAfter}, Column: {colAfter}, " +
                              $"FrozenRows: {frozenRowsAfter}, FrozenColumns: {frozenColsAfter}");

            // Simple validation
            if (isFrozenBefore == isFrozenAfter &&
                rowBefore == rowAfter &&
                colBefore == colAfter &&
                frozenRowsBefore == frozenRowsAfter &&
                frozenColsBefore == frozenColsAfter)
            {
                Console.WriteLine("Freeze state remained unchanged after renaming.");
            }
            else
            {
                Console.WriteLine("Freeze state changed after renaming.");
            }

            // Save the workbook
            string outputPath = "RenameWorksheetAfterFreezeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
