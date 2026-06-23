using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsValidationErrorExport
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook and sample data --------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet dataSheet = workbook.Worksheets[0];          // first worksheet holds source data
            Cells cells = dataSheet.Cells;

            // Header
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Score");

            // Sample rows (some valid, some invalid according to validation rule)
            cells["A2"].PutValue(1);   cells["B2"].PutValue(15);   // valid
            cells["A3"].PutValue(2);   cells["B3"].PutValue(5);    // invalid (below 10)
            cells["A4"].PutValue(3);   cells["B4"].PutValue(25);   // invalid (above 20)
            cells["A5"].PutValue(4);   cells["B5"].PutValue(12);   // valid

            // -------------------- Add data validation (WholeNumber between 10 and 20) --------------------
            Validation validation = dataSheet.Validations[dataSheet.Validations.Add()];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "10";
            validation.Formula2 = "20";
            validation.AddArea(new CellArea { StartRow = 1, StartColumn = 1, EndRow = 4, EndColumn = 1 }); // column B rows 2-5
            validation.ShowError = true; // show error when invalid data entered

            // -------------------- Prepare error report worksheet --------------------
            Worksheet errorSheet = workbook.Worksheets.Add("ErrorReport");
            Cells errorCells = errorSheet.Cells;

            // Copy header to error sheet
            errorCells["A1"].PutValue("ID");
            errorCells["B1"].PutValue("Score");
            errorCells["C1"].PutValue("ErrorMessage");

            int errorRowIndex = 1; // start after header (zero‑based)

            // -------------------- Scan rows and detect validation failures --------------------
            for (int row = 1; row <= 4; row++) // rows 2‑5 in zero‑based index
            {
                // Retrieve the value from the Score column (B)
                object scoreObj = cells[row, 1].Value;
                bool isValid = true;
                string errorMsg = string.Empty;

                // Simple manual validation matching the rule defined above
                if (scoreObj is double || scoreObj is int || scoreObj is decimal)
                {
                    double score = Convert.ToDouble(scoreObj);
                    if (score < 10 || score > 20)
                    {
                        isValid = false;
                        errorMsg = $"Score {score} is outside the allowed range (10‑20).";
                    }
                }
                else
                {
                    isValid = false;
                    errorMsg = "Score is not a numeric value.";
                }

                // If validation failed, copy the entire row to the error sheet
                if (!isValid)
                {
                    // Copy ID
                    errorCells[errorRowIndex, 0].PutValue(cells[row, 0].Value);
                    // Copy Score
                    errorCells[errorRowIndex, 1].PutValue(scoreObj);
                    // Write error description
                    errorCells[errorRowIndex, 2].PutValue(errorMsg);

                    errorRowIndex++;
                }
            }

            // -------------------- Save the workbook --------------------
            workbook.Save("ValidationErrorReport.xlsx");
        }
    }
}