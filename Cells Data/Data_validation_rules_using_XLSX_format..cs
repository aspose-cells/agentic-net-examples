using System;
using Aspose.Cells;

namespace AsposeCellsDataValidationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // -------------------------------------------------
            // 1. Create a new workbook and get the first sheet
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 2. Add a List validation to cell A1 (comma‑separated values)
            // -------------------------------------------------
            Validation listValidation = sheet.Validations[sheet.Validations.Add()];
            listValidation.Type = ValidationType.List;
            listValidation.AlertStyle = ValidationAlertType.Stop;
            listValidation.Formula1 = "Option1,Option2,Option3";
            listValidation.AddArea(new CellArea { StartRow = 0, StartColumn = 0, EndRow = 0, EndColumn = 0 });

            // -------------------------------------------------
            // 3. Add a WholeNumber validation to cell B1 (between 10 and 20)
            // -------------------------------------------------
            Validation wholeNumberValidation = sheet.Validations[sheet.Validations.Add()];
            wholeNumberValidation.Type = ValidationType.WholeNumber;
            wholeNumberValidation.Operator = OperatorType.Between;
            wholeNumberValidation.Formula1 = "10";
            wholeNumberValidation.Formula2 = "20";
            wholeNumberValidation.ShowError = true;
            wholeNumberValidation.ErrorTitle = "Invalid Input";
            wholeNumberValidation.ErrorMessage = "Enter a whole number between 10 and 20.";
            wholeNumberValidation.AddArea(new CellArea { StartRow = 0, StartColumn = 1, EndRow = 0, EndColumn = 1 });

            // -------------------------------------------------
            // 4. Add a Custom validation to cell C1 that references D1
            // -------------------------------------------------
            // First put a sample value in D1 that the formula will use
            sheet.Cells["D1"].PutValue(100);

            Validation customValidation = sheet.Validations[sheet.Validations.Add()];
            customValidation.Type = ValidationType.Custom;
            customValidation.Operator = OperatorType.GreaterOrEqual;
            // Use SetFormula1 to reference D1 (A1‑style, non‑R1C1, non‑local)
            customValidation.SetFormula1("=D1", false, false);
            customValidation.ShowError = true;
            customValidation.ErrorTitle = "Custom Rule Failed";
            customValidation.ErrorMessage = "Value must be greater than or equal to D1.";
            customValidation.AddArea(new CellArea { StartRow = 0, StartColumn = 2, EndRow = 0, EndColumn = 2 });

            // -------------------------------------------------
            // 5. Save the workbook with validations
            // -------------------------------------------------
            string outputPath = "DataValidationDemo.xlsx";
            workbook.Save(outputPath);                               // save

            // -------------------------------------------------
            // 6. Load the workbook with CheckDataValid set to false
            //    (demonstrates the LoadOptions.CheckDataValid property)
            // -------------------------------------------------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CheckDataValid = false;                     // disable data‑validation checking while loading
            Workbook loadedWorkbook = new Workbook(outputPath, loadOptions);

            // Verify that the validation on A1 is still a List type
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Validation loadedValidation = loadedSheet.Validations.GetValidationInCell(0, 0);
            Console.WriteLine($"Loaded validation type in A1: {loadedValidation.Type}");

            // -------------------------------------------------
            // 7. Save the loaded workbook to a new file (optional)
            // -------------------------------------------------
            loadedWorkbook.Save("DataValidationDemo_Loaded.xlsx");
        }
    }
}