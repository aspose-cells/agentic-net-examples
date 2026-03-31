using System;
using Aspose.Cells;

class CheckDropDown
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the cell to check (A1 -> row 0, column 0)
        int row = 0;
        int column = 0;

        // Retrieve the validation applied to the cell, if any
        Validation validation = worksheet.Validations.GetValidationInCell(row, column);

        // Determine whether the validation is a list with an in‑cell drop‑down
        bool isDropDown = false;
        if (validation != null && validation.Type == ValidationType.List)
        {
            isDropDown = validation.InCellDropDown;
        }

        Console.WriteLine($"Cell A1 has drop‑down list: {isDropDown}");
    }
}