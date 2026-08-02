using System;
using Aspose.Cells;

namespace AsposeCellsDataValidationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Import XML data into the first worksheet starting at cell A1 (row 0, column 0)
            // Adjust the file path, sheet name, and start cell as needed
            string xmlPath = "data.xml";
            string targetSheet = "Sheet1";
            workbook.ImportXml(xmlPath, targetSheet, 0, 0);

            // Access the worksheet where XML data was imported
            Worksheet worksheet = workbook.Worksheets[targetSheet];

            // ------------------------------------------------------------
            // Example 1: Whole number validation (e.g., Quantity column)
            // ------------------------------------------------------------
            // Define the range for the validation (e.g., column B, rows 2 to 100)
            CellArea quantityArea = CellArea.CreateCellArea(1, 1, 99, 1); // B2:B100

            // Add a validation to the worksheet for the defined area
            int qtyValidationIndex = worksheet.Validations.Add(quantityArea);
            Validation qtyValidation = worksheet.Validations[qtyValidationIndex];

            // Configure the validation: whole numbers between 1 and 1000
            qtyValidation.Type = ValidationType.WholeNumber;
            qtyValidation.Operator = OperatorType.Between;
            qtyValidation.Formula1 = "1";
            qtyValidation.Formula2 = "1000";
            qtyValidation.InputMessage = "Enter a quantity between 1 and 1000.";
            qtyValidation.ErrorMessage = "Invalid quantity. Must be a whole number between 1 and 1000.";
            qtyValidation.ShowInput = true;
            qtyValidation.ShowError = true;
            qtyValidation.InCellDropDown = false;

            // ------------------------------------------------------------
            // Example 2: List validation (e.g., Status column)
            // ------------------------------------------------------------
            // Define the range for the validation (e.g., column C, rows 2 to 100)
            CellArea statusArea = CellArea.CreateCellArea(1, 2, 99, 2); // C2:C100

            // Add a validation for the status column
            int statusValidationIndex = worksheet.Validations.Add(statusArea);
            Validation statusValidation = worksheet.Validations[statusValidationIndex];

            // Configure the validation: list of allowed status values
            statusValidation.Type = ValidationType.List;
            // Formula1 can be a comma‑separated list or a reference to a range; using a list here
            statusValidation.Formula1 = "\"New,In Progress,Completed,Cancelled\"";
            statusValidation.InCellDropDown = true;
            statusValidation.ShowInput = true;
            statusValidation.InputMessage = "Select a status from the list.";
            statusValidation.ShowError = true;
            statusValidation.ErrorMessage = "Invalid status selected.";

            // ------------------------------------------------------------
            // Example 3: Date validation (e.g., OrderDate column)
            // ------------------------------------------------------------
            // Define the range for the validation (e.g., column D, rows 2 to 100)
            CellArea dateArea = CellArea.CreateCellArea(1, 3, 99, 3); // D2:D100

            int dateValidationIndex = worksheet.Validations.Add(dateArea);
            Validation dateValidation = worksheet.Validations[dateValidationIndex];

            // Configure the validation: dates between today and one year from today
            dateValidation.Type = ValidationType.Date;
            dateValidation.Operator = OperatorType.Between;
            // Use TODAY() function for the start date and TODAY()+365 for the end date
            dateValidation.Formula1 = "=TODAY()";
            dateValidation.Formula2 = "=TODAY()+365";
            dateValidation.InputMessage = "Enter a date within the next year.";
            dateValidation.ErrorMessage = "Date must be between today and one year from today.";
            dateValidation.ShowInput = true;
            dateValidation.ShowError = true;
            dateValidation.InCellDropDown = false;

            // ------------------------------------------------------------
            // Save the workbook with validations applied
            // ------------------------------------------------------------
            workbook.Save("ValidatedOutput.xlsx");
        }
    }
}