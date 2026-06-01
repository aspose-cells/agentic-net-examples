using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class VerifyDeleteOptionsUpdateReference
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data and formulas that reference column A
            cells["A1"].PutValue(10);
            cells["B1"].Formula = "=A1*2";
            cells["A2"].PutValue(20);
            cells["B2"].Formula = "=A2*2";

            // Create DeleteOptions and ensure UpdateReference is true
            DeleteOptions deleteOptions = new DeleteOptions();
            if (!deleteOptions.UpdateReference)
            {
                // Set to true if it is not already set
                deleteOptions.UpdateReference = true;
            }

            // Verify the property before performing deletion
            Console.WriteLine("DeleteOptions.UpdateReference before deletion: " + deleteOptions.UpdateReference);

            // Delete column A (index 0) using DeleteColumns with DeleteOptions
            sheet.Cells.DeleteColumns(0, 1, deleteOptions);

            // After deletion, the original column B becomes column A.
            // Display the updated formula to confirm that references were adjusted.
            Console.WriteLine("Formula in new A1 after column deletion: " + cells["A1"].Formula);

            // Save the modified workbook
            workbook.Save("VerifiedDeleteOptions.xlsx");
        }
    }
}