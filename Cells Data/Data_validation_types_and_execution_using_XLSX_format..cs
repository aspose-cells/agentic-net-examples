using System;
using Aspose.Cells;

namespace AsposeCellsDataValidationDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and get the first worksheet
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 2. Add various data validations
            // -------------------------------------------------

            // Whole number validation (A1) between 10 and 100
            CellArea wholeNumberArea = CellArea.CreateCellArea(0, 0, 0, 0);
            int wholeNumberIndex = sheet.Validations.Add(wholeNumberArea);
            Validation wholeNumberValidation = sheet.Validations[wholeNumberIndex];
            wholeNumberValidation.Type = ValidationType.WholeNumber;
            wholeNumberValidation.Operator = OperatorType.Between;
            wholeNumberValidation.Formula1 = "10";
            wholeNumberValidation.Formula2 = "100";
            wholeNumberValidation.InputMessage = "Enter a whole number between 10 and 100.";
            wholeNumberValidation.ErrorMessage = "Invalid whole number.";
            wholeNumberValidation.ShowInput = true;
            wholeNumberValidation.ShowError = true;

            // Decimal validation (B1) between 0.5 and 9.9
            CellArea decimalArea = CellArea.CreateCellArea(0, 1, 0, 1);
            int decimalIndex = sheet.Validations.Add(decimalArea);
            Validation decimalValidation = sheet.Validations[decimalIndex];
            decimalValidation.Type = ValidationType.Decimal;
            decimalValidation.Operator = OperatorType.Between;
            decimalValidation.Formula1 = "0.5";
            decimalValidation.Formula2 = "9.9";
            decimalValidation.InputMessage = "Enter a decimal between 0.5 and 9.9.";
            decimalValidation.ErrorMessage = "Invalid decimal value.";
            decimalValidation.ShowInput = true;
            decimalValidation.ShowError = true;

            // List validation (C1) with comma‑separated values
            CellArea listArea = CellArea.CreateCellArea(0, 2, 0, 2);
            int listIndex = sheet.Validations.Add(listArea);
            Validation listValidation = sheet.Validations[listIndex];
            listValidation.Type = ValidationType.List;
            listValidation.Formula1 = "Red,Green,Blue";
            listValidation.InCellDropDown = true;
            listValidation.InputMessage = "Select a color from the list.";
            listValidation.ErrorMessage = "Invalid selection.";
            listValidation.ShowInput = true;
            listValidation.ShowError = true;

            // Date validation (D1) between 2023‑01‑01 and 2023‑12‑31
            CellArea dateArea = CellArea.CreateCellArea(0, 3, 0, 3);
            int dateIndex = sheet.Validations.Add(dateArea);
            Validation dateValidation = sheet.Validations[dateIndex];
            dateValidation.Type = ValidationType.Date;
            dateValidation.Operator = OperatorType.Between;
            dateValidation.Formula1 = "DATE(2023,1,1)";
            dateValidation.Formula2 = "DATE(2023,12,31)";
            dateValidation.InputMessage = "Enter a date in 2023.";
            dateValidation.ErrorMessage = "Date out of range.";
            dateValidation.ShowInput = true;
            dateValidation.ShowError = true;

            // Time validation (E1) between 09:00 and 17:00
            CellArea timeArea = CellArea.CreateCellArea(0, 4, 0, 4);
            int timeIndex = sheet.Validations.Add(timeArea);
            Validation timeValidation = sheet.Validations[timeIndex];
            timeValidation.Type = ValidationType.Time;
            timeValidation.Operator = OperatorType.Between;
            timeValidation.Formula1 = "TIME(9,0,0)";
            timeValidation.Formula2 = "TIME(17,0,0)";
            timeValidation.InputMessage = "Enter a time between 09:00 and 17:00.";
            timeValidation.ErrorMessage = "Time out of range.";
            timeValidation.ShowInput = true;
            timeValidation.ShowError = true;

            // Text length validation (F1) between 5 and 10 characters
            CellArea textLengthArea = CellArea.CreateCellArea(0, 5, 0, 5);
            int textLengthIndex = sheet.Validations.Add(textLengthArea);
            Validation textLengthValidation = sheet.Validations[textLengthIndex];
            textLengthValidation.Type = ValidationType.TextLength;
            textLengthValidation.Operator = OperatorType.Between;
            textLengthValidation.Formula1 = "5";
            textLengthValidation.Formula2 = "10";
            textLengthValidation.InputMessage = "Enter text 5‑10 characters long.";
            textLengthValidation.ErrorMessage = "Invalid text length.";
            textLengthValidation.ShowInput = true;
            textLengthValidation.ShowError = true;

            // Custom validation (G1) – value must be double of A1
            CellArea customArea = CellArea.CreateCellArea(0, 6, 0, 6);
            int customIndex = sheet.Validations.Add(customArea);
            Validation customValidation = sheet.Validations[customIndex];
            customValidation.Type = ValidationType.Custom;
            customValidation.Formula1 = "=A1*2";
            customValidation.InputMessage = "Value must be double of A1.";
            customValidation.ErrorMessage = "Custom validation failed.";
            customValidation.ShowInput = true;
            customValidation.ShowError = true;

            // -------------------------------------------------
            // 3. Save the workbook with all validations
            // -------------------------------------------------
            string outputPath = "DataValidationDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");

            // -------------------------------------------------
            // 4. Load the workbook with CheckDataValid set to false
            //    (this skips validation of the template while loading)
            // -------------------------------------------------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CheckDataValid = false;
            Workbook loadedWorkbook = new Workbook(outputPath, loadOptions);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // -------------------------------------------------
            // 5. Retrieve a validation and display its Formula1
            // -------------------------------------------------
            Validation loadedWholeNumberValidation = loadedSheet.Validations.GetValidationInCell(0, 0);
            string formulaA1 = loadedWholeNumberValidation.GetFormula1(false, false);
            string formulaR1C1 = loadedWholeNumberValidation.GetFormula1(true, false);
            Console.WriteLine($"WholeNumber Validation Formula (A1 notation): {formulaA1}");
            Console.WriteLine($"WholeNumber Validation Formula (R1C1 notation): {formulaR1C1}");
        }
    }
}