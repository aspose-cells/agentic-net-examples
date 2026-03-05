using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DataValidationTypesDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Whole number validation (between 10 and 100) on cell A1
            Validation wholeNumber = sheet.Validations[sheet.Validations.Add()];
            wholeNumber.Type = ValidationType.WholeNumber;
            wholeNumber.Operator = OperatorType.Between;
            wholeNumber.Formula1 = "10";
            wholeNumber.Formula2 = "100";
            wholeNumber.InputMessage = "Enter a whole number between 10 and 100.";
            wholeNumber.ErrorMessage = "Invalid whole number.";
            wholeNumber.AddArea(new CellArea { StartRow = 0, StartColumn = 0, EndRow = 0, EndColumn = 0 });

            // Decimal validation (greater than 0.5) on cell B1
            Validation decimalVal = sheet.Validations[sheet.Validations.Add()];
            decimalVal.Type = ValidationType.Decimal;
            decimalVal.Operator = OperatorType.GreaterThan;
            decimalVal.Formula1 = "0.5";
            decimalVal.InputMessage = "Enter a decimal greater than 0.5.";
            decimalVal.ErrorMessage = "Invalid decimal.";
            decimalVal.AddArea(new CellArea { StartRow = 0, StartColumn = 1, EndRow = 0, EndColumn = 1 });

            // List validation (static list) on cell C1
            Validation listVal = sheet.Validations[sheet.Validations.Add()];
            listVal.Type = ValidationType.List;
            listVal.InCellDropDown = true;
            listVal.Formula1 = "Red,Green,Blue";
            listVal.InputMessage = "Select a color.";
            listVal.ErrorMessage = "Invalid selection.";
            listVal.AddArea(new CellArea { StartRow = 0, StartColumn = 2, EndRow = 0, EndColumn = 2 });

            // Date validation (between two dates) on cell D1
            Validation dateVal = sheet.Validations[sheet.Validations.Add()];
            dateVal.Type = ValidationType.Date;
            dateVal.Operator = OperatorType.Between;
            dateVal.Formula1 = "DATE(2023,1,1)";
            dateVal.Formula2 = "DATE(2023,12,31)";
            dateVal.InputMessage = "Enter a date in 2023.";
            dateVal.ErrorMessage = "Date out of range.";
            dateVal.AddArea(new CellArea { StartRow = 0, StartColumn = 3, EndRow = 0, EndColumn = 3 });

            // Time validation (less than 18:00) on cell E1
            Validation timeVal = sheet.Validations[sheet.Validations.Add()];
            timeVal.Type = ValidationType.Time;
            timeVal.Operator = OperatorType.LessThan;
            timeVal.Formula1 = "TIME(18,0,0)";
            timeVal.InputMessage = "Enter a time before 18:00.";
            timeVal.ErrorMessage = "Time must be before 18:00.";
            timeVal.AddArea(new CellArea { StartRow = 0, StartColumn = 4, EndRow = 0, EndColumn = 4 });

            // Text length validation (max 5 characters) on cell F1
            Validation textLen = sheet.Validations[sheet.Validations.Add()];
            textLen.Type = ValidationType.TextLength;
            textLen.Operator = OperatorType.LessThanOrEqual;
            textLen.Formula1 = "5";
            textLen.InputMessage = "Enter up to 5 characters.";
            textLen.ErrorMessage = "Text too long.";
            textLen.AddArea(new CellArea { StartRow = 0, StartColumn = 5, EndRow = 0, EndColumn = 5 });

            // Custom validation (value must be double of A1) on cell G1
            Validation customVal = sheet.Validations[sheet.Validations.Add()];
            customVal.Type = ValidationType.Custom;
            customVal.Formula1 = "=A1*2";
            customVal.InputMessage = "Value must be double of A1.";
            customVal.ErrorMessage = "Invalid custom rule.";
            customVal.AddArea(new CellArea { StartRow = 0, StartColumn = 6, EndRow = 0, EndColumn = 6 });

            // Any value validation (no restriction) on cell H1
            Validation anyVal = sheet.Validations[sheet.Validations.Add()];
            anyVal.Type = ValidationType.AnyValue;
            anyVal.AddArea(new CellArea { StartRow = 0, StartColumn = 7, EndRow = 0, EndColumn = 7 });

            // Save the workbook as XLSX
            workbook.Save("DataValidationTypes.xlsx");
        }
    }
}