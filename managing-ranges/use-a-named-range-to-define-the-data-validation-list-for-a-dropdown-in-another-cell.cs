using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeValidationDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells A1:A5 with the list items
            string[] items = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
            for (int i = 0; i < items.Length; i++)
            {
                worksheet.Cells[i, 0].PutValue(items[i]); // Column A (index 0)
            }

            // Define a named range "FruitList" that refers to A1:A5
            int nameIndex = workbook.Worksheets.Names.Add("FruitList");
            // The RefersTo string must start with '=' and use absolute references
            workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$A$5";

            // Create a validation for cell B1 (row 0, column 1)
            CellArea validationArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 1,
                EndRow = 0,
                EndColumn = 1
            };
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Set validation type to List and point to the named range
            validation.Type = ValidationType.List;
            // When using a named range, specify the name without the leading '='
            validation.Formula1 = "FruitList";
            // Enable the in‑cell drop‑down arrow
            validation.InCellDropDown = true;

            // Save the workbook
            workbook.Save("NamedRangeDropDownDemo.xlsx");
        }
    }
}