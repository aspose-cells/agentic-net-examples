using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some cells with values and a formula that references column A
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["B1"].PutValue(20);
        worksheet.Cells["C1"].Formula = "=A1+B1";

        // Create DeleteOptions and enable reference updating
        DeleteOptions deleteOptions = new DeleteOptions
        {
            UpdateReference = true // Enable formula reference updates after deletion
        };

        // Delete column A (index 0) using the options with UpdateReference = true
        worksheet.Cells.DeleteColumns(0, 1, deleteOptions);

        // Save the modified workbook
        workbook.Save("DeleteColumnWithUpdateReference.xlsx");
    }
}