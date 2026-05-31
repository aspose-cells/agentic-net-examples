using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsValidationErrorReport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a validation for cells A1:A5 (whole numbers between 10 and 20)
            CellArea area = CellArea.CreateCellArea("A1", "A5");
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "10";
            validation.Formula2 = "20";
            validation.AlertStyle = ValidationAlertType.Stop;
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "Value must be between 10 and 20.";
            validation.ShowError = true;

            // Insert some test values (some valid, some invalid)
            sheet.Cells["A1"].PutValue(5);   // Invalid
            sheet.Cells["A2"].PutValue(15);  // Valid
            sheet.Cells["A3"].PutValue(25);  // Invalid
            sheet.Cells["A4"].PutValue(12);  // Valid
            sheet.Cells["A5"].PutValue(8);   // Invalid

            // Prepare a text file to store validation error details
            string errorReportPath = "ValidationErrors.txt";
            using (StreamWriter writer = new StreamWriter(errorReportPath, false))
            {
                // Iterate through all validations in the worksheet
                foreach (Validation val in sheet.Validations)
                {
                    // For each area covered by the validation, check each cell
                    foreach (CellArea valArea in val.Areas)
                    {
                        for (int row = valArea.StartRow; row <= valArea.EndRow; row++)
                        {
                            for (int col = valArea.StartColumn; col <= valArea.EndColumn; col++)
                            {
                                Cell cell = sheet.Cells[row, col];
                                // If the cell value violates the validation, write details to the file
                                if (!IsCellValueValid(cell, val))
                                {
                                    string cellName = CellsHelper.CellIndexToName(row, col);
                                    writer.WriteLine($"Cell {cellName}: {val.ErrorMessage}");
                                }
                            }
                        }
                    }
                }
            }

            // Save the workbook (optional, just to keep the file)
            workbook.Save("ValidationDemo.xlsx");
        }

        // Helper method to evaluate whether a cell satisfies a given validation
        private static bool IsCellValueValid(Cell cell, Validation validation)
        {
            // Use the built‑in validation check by attempting to apply the rule.
            // Aspose.Cells does not expose a direct method, so we perform a simple check
            // for WholeNumber between two values as an example.
            if (validation.Type == ValidationType.WholeNumber && validation.Operator == OperatorType.Between)
            {
                if (double.TryParse(cell.StringValue, out double numericValue))
                {
                    if (double.TryParse(validation.Formula1, out double lower) &&
                        double.TryParse(validation.Formula2, out double upper))
                    {
                        return numericValue >= lower && numericValue <= upper;
                    }
                }
                // Non‑numeric or out of range values are invalid
                return false;
            }

            // For other validation types, assume valid (extend as needed)
            return true;
        }
    }
}