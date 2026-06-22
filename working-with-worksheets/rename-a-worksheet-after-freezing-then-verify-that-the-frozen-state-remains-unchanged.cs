using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RenameWorksheetAfterFreezeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Freeze panes at cell C3 with 3 frozen rows and 3 frozen columns
                sheet.FreezePanes("C3", 3, 3);

                // Capture the freeze state before renaming
                int rowBefore, colBefore, frozenRowsBefore, frozenColsBefore;
                bool isFrozenBefore = sheet.GetFreezedPanes(out rowBefore, out colBefore, out frozenRowsBefore, out frozenColsBefore);
                Console.WriteLine($"Before rename - IsFrozen: {isFrozenBefore}, Row: {rowBefore}, Column: {colBefore}, FrozenRows: {frozenRowsBefore}, FrozenCols: {frozenColsBefore}");

                // Rename the worksheet
                string originalName = sheet.Name;
                string newName = "RenamedSheet";
                sheet.Name = newName;
                Console.WriteLine($"Worksheet renamed from '{originalName}' to '{sheet.Name}'");

                // Verify that the freeze state is unchanged after renaming
                int rowAfter, colAfter, frozenRowsAfter, frozenColsAfter;
                bool isFrozenAfter = sheet.GetFreezedPanes(out rowAfter, out colAfter, out frozenRowsAfter, out frozenColsAfter);
                Console.WriteLine($"After rename - IsFrozen: {isFrozenAfter}, Row: {rowAfter}, Column: {colAfter}, FrozenRows: {frozenRowsAfter}, FrozenCols: {frozenColsAfter}");

                // Simple validation
                bool freezeStateUnchanged = isFrozenBefore == isFrozenAfter &&
                                            rowBefore == rowAfter &&
                                            colBefore == colAfter &&
                                            frozenRowsBefore == frozenRowsAfter &&
                                            frozenColsBefore == frozenColsAfter;
                Console.WriteLine($"Freeze state unchanged: {freezeStateUnchanged}");

                // Save the workbook
                string outputPath = "RenameWorksheetAfterFreezeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
            RenameWorksheetAfterFreezeDemo.Run();
        }
    }
}