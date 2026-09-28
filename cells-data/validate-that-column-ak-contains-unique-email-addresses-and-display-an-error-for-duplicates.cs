// Title: Validate unique email addresses in column AK using Aspose.Cells C# custom data validation
// AI Prompts: Write C# code with Aspose.Cells that adds a custom ValidationType.Custom rule using a COUNTIF formula to guarantee each email in column AK appears only once and shows an error for duplicates. | Generate a snippet that creates a CellArea for column AK, sets the validation formula =COUNTIF($AK:$AK,AK2)=1, and configures an error title and message for duplicate email entries. | Provide a complete example that loads an Excel workbook, applies the uniqueness validation to column AK, and saves the workbook with the rule applied.
// Common Searches: aspocells c# validate unique emails in a specific column | how to add a custom COUNTIF validation to an Excel column using Aspose.Cells .NET | prevent duplicate email entries in Excel with Aspose.Cells validation rule | set error message for duplicate values in column AK programmatically Aspose.Cells | c# aspocells data validation for email uniqueness
// Tags: Aspose.Cells custom validation COUNTIF | C# email uniqueness validation Excel | Aspose.Cells column AK data validation | Excel duplicate email error message .NET | programmatic data validation Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsEmailValidation
{
    // The example loads an existing workbook, defines a CellArea covering column AK, adds a custom validation with a COUNTIF formula to ensure each email address is unique, configures a duplicate‑email error title and message, and saves the workbook with the validation applied.
    class Program
    {
        // Convert zero‑based column index to Excel column name (e.g., 0 -> "A")
        static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = String.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Column AK is the 37th column (zero‑based index 36)
            int emailColumnIndex = 36;

            // Determine the last row that contains data
            int lastDataRow = worksheet.Cells.MaxDataRow; // zero‑based

            // Assume the first row (row 0) contains a header, start validation from row 1
            int startRow = 1;

            // Define the area for the validation (entire column AK from startRow to lastDataRow)
            CellArea emailArea = new CellArea
            {
                StartRow = startRow,
                EndRow = lastDataRow,
                StartColumn = emailColumnIndex,
                EndColumn = emailColumnIndex
            };

            // Add a custom validation to the defined area
            int validationIndex = worksheet.Validations.Add(emailArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Set validation type to Custom
            validation.Type = ValidationType.Custom;

            // Build the formula: =COUNTIF($AK:$AK,AK2)=1
            // $AK:$AK locks the whole column, AK2 is a relative reference to the top‑left cell of the area
            string columnLetter = GetColumnName(emailColumnIndex);
            string formula = $"=COUNTIF(${columnLetter}:${columnLetter},{columnLetter}{startRow + 1})=1";
            validation.Formula1 = formula;

            // Configure the error message that will be shown when a duplicate is entered
            validation.ShowError = true;
            validation.ErrorTitle = "Duplicate Email";
            validation.ErrorMessage = "Each email address in column AK must be unique.";

            // Save the workbook with the validation applied
            workbook.Save("output.xlsx");
        }
    }
}
