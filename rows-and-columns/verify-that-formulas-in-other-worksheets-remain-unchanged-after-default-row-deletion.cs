// Title: Check that a cross‑sheet formula remains unchanged after deleting a row with Aspose.Cells in C#
// AI Prompts: Write C# code using Aspose.Cells to create two worksheets, set a formula in one sheet that references a cell in the other, delete a specific row in the source sheet with DeleteRow, and confirm the formula text does not change. | Demonstrate how to recalculate the workbook after the row deletion and output the evaluated value of the cross‑sheet formula to verify it still points to the original cell.
// Common Searches: Aspose.Cells DeleteRow does not adjust formulas in other worksheets C# | how to keep external cell reference unchanged after row removal Aspose.Cells | verify formula reference stability after deleting rows in a workbook using Aspose.Cells | C# Aspose.Cells cross‑sheet formula behavior when source row is deleted | default DeleteRow method external reference handling Aspose.Cells example
// Tags: DeleteRow without updating external references Aspose.Cells | C# verify formula reference after row removal | Aspose.Cells workbook multiple sheets formula integrity | preserve cross-sheet cell reference Aspose.Cells | default DeleteRow external reference handling

using System;
using Aspose.Cells;

namespace VerifyFormulaUnchangedAfterRowDeletion
{
    // Shows how to create a workbook with two worksheets, add a cross‑sheet formula, delete a row in the source sheet using DeleteRow, and confirm that the formula on the other sheet stays unchanged after recalculation.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Add a second worksheet for formulas that reference the first sheet
            Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
            summarySheet.Name = "Summary";

            // Populate the Data sheet with sample values (rows 0‑9, column A)
            for (int i = 0; i < 10; i++)
            {
                dataSheet.Cells[i, 0].PutValue(i + 1); // A1 = 1, A2 = 2, ...
            }

            // Set a formula in the Summary sheet that references a cell in the Data sheet
            // Initially points to Data!A5 (row index 4)
            summarySheet.Cells["A1"].Formula = "=Data!A5";

            // Display the original formula
            Console.WriteLine("Original formula in Summary!A1: " + summarySheet.Cells["A1"].Formula);

            // Delete row 4 (zero‑based index 3) from the Data sheet using the default DeleteRow method
            // This method does NOT update references in other worksheets
            dataSheet.Cells.DeleteRow(3);

            // After deletion, the formula in Summary!A1 should remain unchanged
            Console.WriteLine("Formula after deleting row 4 in Data sheet: " + summarySheet.Cells["A1"].Formula);

            // Optionally, verify the value that the formula now evaluates to
            workbook.CalculateFormula();
            Console.WriteLine("Evaluated value of Summary!A1: " + summarySheet.Cells["A1"].StringValue);

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("VerifyFormulaUnchanged.xlsx");
        }
    }
}
