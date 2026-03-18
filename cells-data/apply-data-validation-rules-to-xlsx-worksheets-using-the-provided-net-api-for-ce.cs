using System;
using Aspose.Cells;

namespace AsposeCellsDataValidationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 1. Whole number validation (between 3 and 1234) applied to cell A1
            // ------------------------------------------------------------
            // Define the cell area for A1 (row 0, column 0)
            CellArea wholeNumberArea = CellArea.CreateCellArea(0, 0, 0, 0);
            // Add a new validation to the worksheet for the defined area
            int wholeNumberIndex = worksheet.Validations.Add(wholeNumberArea);
            Validation wholeNumberValidation = worksheet.Validations[wholeNumberIndex];
            // Configure validation properties
            wholeNumberValidation.Type = ValidationType.WholeNumber;
            wholeNumberValidation.Operator = OperatorType.Between;
            wholeNumberValidation.Formula1 = "3";
            wholeNumberValidation.Formula2 = "1234";
            wholeNumberValidation.InputMessage = "Enter a whole number between 3 and 1234.";
            wholeNumberValidation.ErrorMessage = "Invalid entry! Value must be between 3 and 1234.";
            wholeNumberValidation.ShowInput = true;
            wholeNumberValidation.ShowError = true;
            wholeNumberValidation.InCellDropDown = false;

            // ------------------------------------------------------------
            // 2. List validation (drop‑down list) applied to range B2:B5
            // ------------------------------------------------------------
            // Define the cell area for B2:B5 (rows 1‑4, column 1)
            CellArea listArea = CellArea.CreateCellArea(1, 1, 4, 1);
            int listIndex = worksheet.Validations.Add(listArea);
            Validation listValidation = worksheet.Validations[listIndex];
            listValidation.Type = ValidationType.List;
            // Provide a comma‑separated list of allowed values
            listValidation.Formula1 = "Option1,Option2,Option3,Option4";
            listValidation.InCellDropDown = true;
            listValidation.ShowInput = true;
            listValidation.InputMessage = "Select an option from the list.";
            listValidation.ShowError = true;
            listValidation.ErrorMessage = "Please select a valid option.";

            // ------------------------------------------------------------
            // 3. Date validation (between two dates) applied to cell C3
            // ------------------------------------------------------------
            // Define the cell area for C3 (row 2, column 2)
            CellArea dateArea = CellArea.CreateCellArea(2, 2, 2, 2);
            int dateIndex = worksheet.Validations.Add(dateArea);
            Validation dateValidation = worksheet.Validations[dateIndex];
            dateValidation.Type = ValidationType.Date;
            dateValidation.Operator = OperatorType.Between;
            // Use Excel date serial numbers or A1‑style formulas
            dateValidation.Formula1 = "DATE(2023,1,1)";   // start date
            dateValidation.Formula2 = "DATE(2023,12,31)"; // end date
            dateValidation.InputMessage = "Enter a date in 2023.";
            dateValidation.ErrorMessage = "Date must be within the year 2023.";
            dateValidation.ShowInput = true;
            dateValidation.ShowError = true;

            // ------------------------------------------------------------
            // Save the workbook to an XLSX file
            // ------------------------------------------------------------
            workbook.Save("DataValidationDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}