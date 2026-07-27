using System;
using Aspose.Cells;

namespace AsposeCellsValidationExport
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your source file)
            Workbook workbook = new Workbook("SourceData.xlsx");

            // Source worksheet (assumed first sheet)
            Worksheet srcSheet = workbook.Worksheets[0];

            // Create a new worksheet to hold rows that fail validation
            Worksheet errorSheet = workbook.Worksheets.Add("ValidationErrors");

            // Index for the next row to write in the error sheet
            int errorRowIndex = 0;

            // Determine the last row that contains data in the source sheet (using column 0 as reference)
            int lastRow = srcSheet.Cells.GetLastDataRow(0);

            // Iterate through each data row
            for (int row = 0; row <= lastRow; row++)
            {
                bool rowIsValid = true;

                // Check each validation rule defined in the worksheet
                foreach (Validation validation in srcSheet.Validations)
                {
                    // Examine each area covered by the validation
                    foreach (CellArea area in validation.Areas)
                    {
                        // If the current row lies within the validation area
                        if (row >= area.StartRow && row <= area.EndRow)
                        {
                            // Iterate through columns of the area
                            for (int col = area.StartColumn; col <= area.EndColumn; col++)
                            {
                                // Get the cell value
                                Cell cell = srcSheet.Cells[row, col];
                                string cellValue = cell.StringValue?.Trim();

                                // Simple handling for WholeNumber type with Between operator
                                if (validation.Type == ValidationType.WholeNumber &&
                                    validation.Operator == OperatorType.Between)
                                {
                                    if (int.TryParse(cellValue, out int numericValue))
                                    {
                                        if (int.TryParse(validation.Formula1, out int min) &&
                                            int.TryParse(validation.Formula2, out int max))
                                        {
                                            if (numericValue < min || numericValue > max)
                                            {
                                                rowIsValid = false;
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // Non‑numeric value fails WholeNumber validation
                                        rowIsValid = false;
                                        break;
                                    }
                                }
                                // Additional validation types can be added here following the same pattern
                            }
                        }

                        if (!rowIsValid) break;
                    }

                    if (!rowIsValid) break;
                }

                // If the row failed any validation, copy it to the error sheet
                if (!rowIsValid)
                {
                    // Copy the entire row from source to error sheet
                    errorSheet.Cells.CopyRow(srcSheet.Cells, row, errorRowIndex);
                    errorRowIndex++;
                }
            }

            // Save the workbook with the new error sheet
            workbook.Save("SourceData_WithValidationErrors.xlsx");
        }
    }
}