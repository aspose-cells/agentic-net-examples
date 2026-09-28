// Title: How to copy rows that fail a whole-number validation into a separate worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that iterates through a worksheet, evaluates each cell against a numeric range validation, and moves the entire row to a new sheet named "ValidationErrors". | Demonstrate building an error‑report worksheet that lists all rows where the Age column is outside the 10‑20 range by using Aspose.Cells validation APIs.
// Common Searches: Aspose.Cells .NET copy rows with invalid numeric values to another worksheet | C# create validation error report sheet using Aspose.Cells | How to filter rows based on data validation and export them in Aspose.Cells | Export rows that do not meet whole number validation to a new sheet in Aspose.Cells | Generate Excel error log for failed data validation with Aspose.Cells C#
// Tags: export rows with validation errors Aspose.Cells | generate error report worksheet C# | filter worksheet rows by numeric validation Aspose.Cells | create validation error sheet .NET | copy invalid data to new Excel sheet Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsErrorReporting
{
    // The example creates a workbook, adds sample data with an Age column, applies a whole-number validation (10‑20) to that column, scans each data row, and copies any row that violates the rule to a newly added worksheet named "ValidationErrors" before saving the file as ValidationErrorReport.xlsx.
    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook and populate data --------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet dataSheet = workbook.Worksheets[0];          // first worksheet

            // Header
            dataSheet.Cells["A1"].PutValue("Age");
            dataSheet.Cells["B1"].PutValue("Name");

            // Sample data (some values violate the validation rule)
            dataSheet.Cells["A2"].PutValue(15);   // valid
            dataSheet.Cells["B2"].PutValue("John");
            dataSheet.Cells["A3"].PutValue(5);    // invalid (less than 10)
            dataSheet.Cells["B3"].PutValue("Alice");
            dataSheet.Cells["A4"].PutValue(25);   // invalid (greater than 20)
            dataSheet.Cells["B4"].PutValue("Bob");
            dataSheet.Cells["A5"].PutValue(18);   // valid
            dataSheet.Cells["B5"].PutValue("Eve");

            // -------------------- Add data validation (Whole number between 10 and 20) --------------------
            int validationIndex = dataSheet.Validations.Add();
            Validation validation = dataSheet.Validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "10";
            validation.Formula2 = "20";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Age";
            validation.ErrorMessage = "Age must be between 10 and 20.";
            // Apply validation to column A (excluding header)
            validation.AddArea(new CellArea { StartRow = 1, StartColumn = 0, EndRow = dataSheet.Cells.MaxDataRow, EndColumn = 0 });

            // -------------------- Create a worksheet for error reporting --------------------
            Worksheet errorSheet = workbook.Worksheets.Add("ValidationErrors");
            int errorRow = 0; // start writing from the first row in the error sheet

            // Copy header to error sheet
            for (int col = 0; col <= dataSheet.Cells.MaxDataColumn; col++)
            {
                errorSheet.Cells[errorRow, col].Value = dataSheet.Cells[0, col].Value;
            }
            errorRow++; // move to next row after header

            // -------------------- Scan rows and copy those that fail validation --------------------
            for (int row = 1; row <= dataSheet.Cells.MaxDataRow; row++)
            {
                // Retrieve the cell that has validation (column A)
                Cell ageCell = dataSheet.Cells[row, 0];

                // Perform manual validation check (since Aspose.Cells does not expose a direct method)
                bool isValid = true;
                if (ageCell.Type == CellValueType.IsNumeric)
                {
                    double age = ageCell.DoubleValue;
                    if (age < 10 || age > 20)
                        isValid = false;
                }
                else
                {
                    // Non‑numeric values are also considered invalid for this rule
                    isValid = false;
                }

                // If validation fails, copy the entire row to the error sheet
                if (!isValid)
                {
                    for (int col = 0; col <= dataSheet.Cells.MaxDataColumn; col++)
                    {
                        errorSheet.Cells[errorRow, col].Value = dataSheet.Cells[row, col].Value;
                    }
                    errorRow++;
                }
            }

            // -------------------- Save the workbook --------------------
            workbook.Save("ValidationErrorReport.xlsx"); // save workbook
        }
    }
}
