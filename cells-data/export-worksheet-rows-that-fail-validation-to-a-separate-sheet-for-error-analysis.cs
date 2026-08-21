// Title: Export Invalid Rows to a Separate Worksheet with Aspose.Cells for .NET (C#)
// Description: Loads an Excel file, creates an "ErrorRows" sheet, scans each row of the first worksheet for cells that contain Excel error values or breach data‑validation rules, copies the entire offending row to the new sheet while preserving formulas, formats and styles, and saves the result as a new workbook.
// Keywords: Aspose.Cells | C# | .NET | export rows with errors | Excel error cells | data validation failure | copy row Aspose.Cells | save workbook | error analysis worksheet
// Common Searches: Aspose.Cells copy rows with #DIV/0! to another sheet | C# extract rows that fail data validation using Aspose.Cells | How to create an error report sheet in Excel with Aspose.Cells | Export rows containing error values from a workbook in .NET | Separate invalid rows into a new worksheet with Aspose.Cells
// Developer Intent: Identify rows that contain error values or violate validation rules and move them to a dedicated worksheet for review.
// Use Cases: Generate an error‑report sheet for downstream processing. | Isolate rows that break whole‑number validation ranges. | Audit spreadsheets for #VALUE!, #REF!, and other Excel errors. | Maintain original data integrity while segregating invalid entries.
// AI Prompts: Provide C# code using Aspose.Cells to detect ErrorCellValueType cells and copy the whole row to a new worksheet. | Show how to evaluate whole‑number data‑validation (OperatorType.Between) and export rows that fall outside the limits. | Explain how to preserve formulas, cell styles, and conditional formatting when copying rows with validation errors.

using System;
using Aspose.Cells;

namespace AsposeCellsErrorExport
{
    // Loads an Excel file, creates an "ErrorRows" sheet, scans each row of the first worksheet for cells that contain Excel error values or breach data‑validation rules, copies the entire offending row to the new sheet while preserving formulas, formats and styles, and saves the result as a new workbook.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (lifecycle: load)
            Workbook workbook = new Workbook("InputData.xlsx");

            // Access the worksheet that contains the data to be validated
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add a new worksheet that will hold rows with validation errors
            Worksheet errorSheet = workbook.Worksheets.Add("ErrorRows");

            // Determine the used range of the source sheet
            int maxRow = sourceSheet.Cells.MaxDataRow;
            int maxCol = sourceSheet.Cells.MaxDataColumn;

            // Index for the next row to write in the error sheet
            int errorRowIndex = 0;

            // Iterate through each row in the source sheet
            for (int row = 0; row <= maxRow; row++)
            {
                bool rowHasError = false;

                // Scan all cells in the current row
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sourceSheet.Cells[row, col];

                    // If the cell contains an error value (e.g., #DIV/0!, #VALUE!, etc.)
                    // the Value property will be of type ErrorCellValueType
                    if (cell != null && cell.Value is ErrorCellValueType)
                    {
                        rowHasError = true;
                        break;
                    }

                    // Additionally, check if the cell is subject to a data‑validation rule
                    // and whether the current value violates that rule.
                    // Validation.GetValidationInCell returns null when no validation is applied.
                    Validation validation = sourceSheet.Validations.GetValidationInCell(row, col);
                    if (validation != null)
                    {
                        // Perform a simple validation check:
                        // For WholeNumber between two values, ensure the cell value is numeric
                        // and lies within the defined range. Extend this block for other
                        // validation types as needed.
                        if (validation.Type == ValidationType.WholeNumber &&
                            validation.Operator == OperatorType.Between)
                        {
                            double min = double.Parse(validation.Formula1);
                            double max = double.Parse(validation.Formula2);

                            if (double.TryParse(cell.StringValue, out double numericValue))
                            {
                                if (numericValue < min || numericValue > max)
                                {
                                    rowHasError = true;
                                    break;
                                }
                            }
                            else
                            {
                                // Non‑numeric value violates whole‑number validation
                                rowHasError = true;
                                break;
                            }
                        }
                    }
                }

                // If any cell in the row failed validation, copy the entire row to the error sheet
                if (rowHasError)
                {
                    // CopyRow copies data, formulas, formats, etc.
                    sourceSheet.Cells.CopyRow(sourceSheet.Cells, row, errorRowIndex);
                    errorRowIndex++;
                }
            }

            // Save the workbook with the new error sheet (lifecycle: save)
            workbook.Save("OutputWithErrorRows.xlsx");
        }
    }
}
