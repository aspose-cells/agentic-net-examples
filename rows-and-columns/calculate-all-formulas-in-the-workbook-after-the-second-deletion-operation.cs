using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Prepare delete options to update references in other worksheets if needed
        DeleteOptions deleteOptions = new DeleteOptions
        {
            UpdateReference = true
        };

        // ----- First deletion operation -----
        // Example: delete the second column (index 1)
        sheet.Cells.DeleteColumns(1, 1, deleteOptions);

        // ----- Second deletion operation -----
        // Example: delete the second row (index 1)
        sheet.Cells.DeleteRows(1, 1, deleteOptions);

        // After the second deletion, calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}