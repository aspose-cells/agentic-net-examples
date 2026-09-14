// Title: Convert CSV to XLSX and apply whole-number (1-100) validation to column B using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to read a CSV file, convert it to an XLSX workbook, and then define a validation rule that only permits whole numbers from 1 to 100 in column B. | Create a validation entry for column B of the generated Excel file that enforces a numeric range without showing a dropdown, using the ValidationCollection API in C#. | Persist the workbook after adding the data‑validation rule so the CSV‑to‑XLSX conversion includes the required constraints.
// Common Searches: aspnet convert csv to xlsx and apply validation to column B | c# aspose.cells restrict column B values to 1-100 after csv import | how to enforce whole number limits in an Excel sheet generated from CSV using Aspose.Cells | add numeric constraint to a column in a workbook created from CSV with Aspose.Cells .NET
// Tags: csv to xlsx conversion aspose.cells | range validation aspose.cells | validationcollection usage c# | column b data validation aspose.cells | save workbook with validation aspose.cells

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsDataValidationExample
{
    // The example converts input.csv to output.xlsx with ConversionUtility, loads the workbook, defines a validation area covering all rows in column B starting at row 2, adds a whole-number validation (1‑100) with custom input and error messages, and saves the workbook, ensuring the CSV‑to‑XLSX conversion includes the required data‑validation rule.
    class Program
    {
        static void Main()
        {
            // Paths for source CSV and destination XLSX files
            string csvFile = "input.csv";
            string xlsxFile = "output.xlsx";

            // -------------------------------------------------
            // 1. Convert CSV to XLSX using the provided utility
            // -------------------------------------------------
            ConversionUtility.Convert(csvFile, xlsxFile);

            // -------------------------------------------------
            // 2. Load the newly created workbook
            // -------------------------------------------------
            Workbook workbook = new Workbook(xlsxFile);
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 3. Add a data validation rule to restrict values
            //    in column B (index 1) to whole numbers between 1 and 100
            // -------------------------------------------------
            ValidationCollection validations = worksheet.Validations;

            // Define the range for the validation (all rows in column B, starting from row 2)
            CellArea area = new CellArea
            {
                StartRow = 1,          // Row index 1 (second row, assuming first row is header)
                EndRow = worksheet.Cells.MaxDataRow, // Use the last used row; alternatively use 1048575 for full column
                StartColumn = 1,       // Column B (zero‑based index)
                EndColumn = 1
            };

            // Add the validation to the worksheet
            int validationIndex = validations.Add(area);
            Validation validation = validations[validationIndex];

            // Configure the validation properties
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";   // Minimum allowed value
            validation.Formula2 = "100"; // Maximum allowed value
            validation.ShowInput = true;
            validation.ShowError = true;
            validation.InputMessage = "Please enter a whole number between 1 and 100.";
            validation.ErrorMessage = "Invalid entry. Value must be a whole number between 1 and 100.";
            validation.InCellDropDown = false; // No dropdown needed for numeric range

            // -------------------------------------------------
            // 4. Save the workbook with the validation applied
            // -------------------------------------------------
            workbook.Save(xlsxFile, SaveFormat.Xlsx);

            Console.WriteLine("CSV converted to XLSX and data validation added successfully.");
        }
    }
}
