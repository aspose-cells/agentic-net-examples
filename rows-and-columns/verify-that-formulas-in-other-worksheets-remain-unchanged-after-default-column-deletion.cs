// Title: Aspose.Cells for .NET – Preserve cross‑worksheet formulas when deleting a column
// Description: This C# example creates a workbook with two sheets, writes a value to Sheet1!A1, adds a formula in Sheet2!B1 that points to Sheet1!A1, deletes the first column of Sheet1 using DeleteColumn(0) (default false), and shows that the formula in Sheet2 remains "=Sheet1!A1". The workbook is then saved.
// Keywords: Aspose.Cells DeleteColumn | C# keep formula after column removal | cross sheet reference unchanged | default false DeleteColumn behavior | Aspose.Cells .NET example | preserve external formula references | column deletion without updating formulas
// Common Searches: Aspose.Cells delete column without changing formulas | Does DeleteColumn affect formulas in other worksheets | C# keep Sheet2 formula after removing column in Sheet1 | How to prevent formula updates when deleting a column in Aspose.Cells | Aspose.Cells default DeleteColumn behavior
// Developer Intent: Confirm that a formula on a different worksheet is not altered when the referenced column is removed from the source sheet.
// Use Cases: Validate data‑migration scripts that modify sheet structure but must retain original references. | Run automated checks before publishing a workbook to ensure reporting formulas stay intact. | Demonstrate how to delete columns while preserving formula integrity for financial models.
// AI Prompts: Write C# code with Aspose.Cells that deletes column A in Sheet1 but leaves formulas in other sheets unchanged. | Create a unit test in C# asserting that Sheet2!B1 still equals "=Sheet1!A1" after calling DeleteColumn on Sheet1. | Explain the impact of the DeleteColumn overload with the boolean flag on formula updates in Aspose.Cells.

using System;
using Aspose.Cells;

namespace VerifyFormulaAfterColumnDeletion
{
    // This C# example creates a workbook with two sheets, writes a value to Sheet1!A1, adds a formula in Sheet2!B1 that points to Sheet1!A1, deletes the first column of Sheet1 using DeleteColumn(0) (default false), and shows that the formula in Sheet2 remains "=Sheet1!A1". The workbook is then saved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];               // First sheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");    // Second sheet

            // Populate Sheet1 with a value in cell A1
            sheet1.Cells["A1"].PutValue(100);

            // In Sheet2, set a formula that references Sheet1!A1
            // This formula should remain unchanged after deleting a column in Sheet1
            sheet2.Cells["B1"].Formula = "=Sheet1!A1";

            // Display the original formula in Sheet2
            Console.WriteLine("Original formula in Sheet2!B1: " + sheet2.Cells["B1"].Formula);

            // Delete the first column (A) in Sheet1 without updating references
            // Using the overload without the bool parameter (default is false)
            sheet1.Cells.DeleteColumn(0);

            // After deletion, the formula in Sheet2 should still be "=Sheet1!A1"
            Console.WriteLine("Formula in Sheet2!B1 after deleting column in Sheet1: " + sheet2.Cells["B1"].Formula);

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("VerifyFormulaAfterColumnDeletion.xlsx");
        }
    }
}
