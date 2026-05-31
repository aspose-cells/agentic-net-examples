using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class DeleteOptionsUpdateReferenceDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data and formulas
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);
            cells["C1"].Formula = "=A1+B1";

            cells["A2"].PutValue(30);
            cells["B2"].PutValue(40);
            cells["C2"].Formula = "=A2+B2";

            // Create DeleteOptions with UpdateReference set to false
            // This preserves the original formulas after deletion
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = false
            };

            // Delete column A (index 0) using the options
            // Formulas in C1 and C2 will still reference the original cells A1/B1 and A2/B2
            sheet.Cells.DeleteColumns(0, 1, deleteOptions);

            // Save the workbook
            workbook.Save("DeleteOptionsUpdateReferenceDemo.xlsx");
        }
    }
}