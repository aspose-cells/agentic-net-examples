using System;
using Aspose.Cells;

namespace AsposeCellsInCellDropdownCheck
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or specify the required one)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the cell K10
            Cell cell = worksheet.Cells["K10"];

            // Retrieve the validation applied to the cell
            Validation validation = cell.GetValidation();

            // Determine if the validation exists and if it shows an in‑cell dropdown
            bool hasInCellDropdown = validation != null && validation.InCellDropDown;

            // Output the result to the console
            Console.WriteLine($"Cell K10 uses an in‑cell dropdown: {hasInCellDropdown}");
        }
    }
}