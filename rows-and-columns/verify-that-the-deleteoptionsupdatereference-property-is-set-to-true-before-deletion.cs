// Title: C# – Verify DeleteOptions.UpdateReference before Deleting Columns with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, ensure DeleteOptions.UpdateReference is enabled, delete a column using DeleteColumns, and confirm that formulas adjust automatically before saving the file.
// Keywords: Aspose.Cells | DeleteOptions | UpdateReference | C# | .NET | DeleteColumns | formula reference update | column deletion | workbook manipulation
// Common Searches: Aspose.Cells set DeleteOptions.UpdateReference true | C# delete column keep formulas updated Aspose.Cells | how to verify DeleteOptions before column removal | DeleteColumns with reference update .NET | Aspose.Cells DeleteOptions property check
// Developer Intent: Make sure the UpdateReference flag is true so that deleting columns automatically updates dependent formulas.
// Use Cases: Programmatically delete one or more columns while preserving formula integrity. | Validate and correct DeleteOptions settings before modifying worksheet structure. | Automate workbook cleanup tasks that require reference‑aware column removal.
// AI Prompts: Generate C# code that checks DeleteOptions.UpdateReference, sets it if needed, deletes a range of columns, and returns the modified workbook. | Explain the impact of DeleteOptions.UpdateReference on formula recalculation when rows or columns are removed in Aspose.Cells. | Create a method to delete a column in a .NET workbook, ensuring formulas are updated and output the new formula strings.

using System;
using Aspose.Cells;

namespace DeleteOptionsVerificationDemo
{
    // Demonstrates how to create a workbook, ensure DeleteOptions.UpdateReference is enabled, delete a column using DeleteColumns, and confirm that formulas adjust automatically before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data and a formula that references column A
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);
            cells["C1"].Formula = "=A1+B1";

            // Create DeleteOptions instance
            DeleteOptions deleteOptions = new DeleteOptions();

            // Verify that UpdateReference is true; if not, set it
            if (!deleteOptions.UpdateReference)
            {
                // Ensure references are updated when deleting
                deleteOptions.UpdateReference = true;
                Console.WriteLine("UpdateReference was false; set to true.");
            }
            else
            {
                Console.WriteLine("UpdateReference is already true.");
            }

            // Delete column A (index 0) using the verified DeleteOptions
            // This will shift columns left and update the formula in C1 accordingly
            sheet.Cells.DeleteColumns(0, 1, deleteOptions);

            // Output the updated formula to demonstrate that references were updated
            Console.WriteLine("Updated formula in C1 after column deletion: " + cells["C1"].Formula);

            // Save the workbook
            workbook.Save("DeleteOptionsVerificationOutput.xlsx");
        }
    }
}
