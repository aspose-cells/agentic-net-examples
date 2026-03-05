using System;
using Aspose.Cells;

class CheckDropdownDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the cell to inspect (e.g., A1 -> row 0, column 0)
        int row = 0;
        int column = 0;

        // Retrieve the validation applied to the specified cell
        Validation validation = worksheet.Validations.GetValidationInCell(row, column);

        // Check whether the validation exists and if it displays an in‑cell dropdown
        bool isDropdown = validation != null && validation.InCellDropDown;

        Console.WriteLine($"Cell A1 has in‑cell dropdown: {isDropdown}");

        // Save the workbook (no modifications made, just demonstrating lifecycle usage)
        workbook.Save("output.xlsx");
    }
}