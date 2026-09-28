// Title: Delete columns B and C while automatically updating formulas with DeleteOptions in Aspose.Cells for .NET
// AI Prompts: Instantiate DeleteOptions with UpdateReference = true, call the DeleteColumns overload to remove columns B and C, and verify that the SUM formula shifts to the new range. | Show how to delete multiple adjacent columns in a C# Aspose.Cells workbook and keep all dependent formulas correct by configuring DeleteOptions.
// Common Searches: Aspose.Cells DeleteOptions keep formulas correct C# example | how to delete specific columns and keep formulas updated in Aspose.Cells | C# delete columns B and C and adjust SUM formula Aspose.Cells | using DeleteColumns overload with DeleteOptions in Aspose.Cells .NET | preserve cell references after column deletion Aspose.Cells
// Tags: DeleteOptions reference update | delete columns with formula adjustment Aspose.Cells | Aspose.Cells DeleteColumns overload | C# Excel column removal preserving formulas | Aspose.Cells workbook column deletion example

using System;
using Aspose.Cells;

namespace AsposeCellsDeleteOptionsDemo
{
    // The sample creates a workbook, fills cells A1:D2, adds a SUM formula in E1, configures DeleteOptions with UpdateReference=true, deletes columns B and C via the DeleteColumns overload, automatically updates the formula to reference the new range, prints the updated formula, and saves the file as DeleteOptionsDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data in columns A to D
            cells["A1"].PutValue("Header A");
            cells["B1"].PutValue("Header B");
            cells["C1"].PutValue("Header C");
            cells["D1"].PutValue("Header D");

            cells["A2"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["C2"].PutValue(30);
            cells["D2"].PutValue(40);

            // Add a formula that references the columns we will delete
            cells["E1"].Formula = "=SUM(A2:D2)";

            // Create DeleteOptions and set UpdateReference to true
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = true // ensures formulas adjust after deletion
            };

            // Delete columns B and C (indexes 1 and 2) using the DeleteOptions overload
            // This will shift remaining columns left and update the formula in E1
            sheet.Cells.DeleteColumns(1, 2, deleteOptions);

            // After deletion, column D becomes column B, and the formula should be updated to "=SUM(A2:B2)"
            Console.WriteLine("Formula after deleting columns B and C: " + cells["E1"].Formula);

            // Save the workbook to demonstrate that changes are persisted
            workbook.Save("DeleteOptionsDemo.xlsx");
        }
    }
}
