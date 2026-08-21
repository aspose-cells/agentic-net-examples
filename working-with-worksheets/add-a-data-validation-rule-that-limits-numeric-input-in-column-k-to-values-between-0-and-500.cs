// Title: C# – Add numeric data validation (0‑500) to column K with Aspose.Cells
// Description: Creates a new workbook, defines a CellArea for column K (rows 0‑1000), adds a WholeNumber validation with a Between operator, sets the allowed range to 0‑500, configures custom input and error messages, applies a Stop alert style, and saves the file as ColumnKValidation.xlsx.
// Keywords: Aspose.Cells C# | Excel data validation | numeric range validation | column K validation | whole number 0 to 500 | ValidationType.WholeNumber | OperatorType.Between | validation alert stop | CellArea example | Aspose.Cells workbook
// Common Searches: Aspose.Cells set numeric range validation | C# restrict Excel column values 0 to 500 | Add data validation to column K using Aspose | Aspose.Cells custom input and error messages | Apply whole number validation in .NET Excel
// Developer Intent: Add a rule that allows only whole numbers between 0 and 500 in column K of an Excel worksheet.
// Use Cases: Enforce price limits in a generated financial report. | Prevent out‑of‑range quantities in a user‑filled inventory template. | Guide data entry with custom messages in a survey worksheet. | Ensure data quality in automated Excel exports.
// AI Prompts: Generate Aspose.Cells C# code to apply a whole‑number validation (0‑500) to column K rows 1‑1000 with custom input and error messages. | Show how to modify the rule to accept decimal values between 0 and 500 using Aspose.Cells. | Provide a C# loop that applies the same numeric range validation to columns K, L, and M.

using Aspose.Cells;

// Creates a new workbook, defines a CellArea for column K (rows 0‑1000), adds a WholeNumber validation with a Between operator, sets the allowed range to 0‑500, configures custom input and error messages, applies a Stop alert style, and saves the file as ColumnKValidation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the validation area for column K (zero‑based index 10), rows 0‑1000
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 1000,
            StartColumn = 10,
            EndColumn = 10
        };

        // Add a validation to the worksheet for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Configure the validation: whole numbers between 0 and 500
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "0";
        validation.Formula2 = "500";

        // Optional user‑friendly messages and alert style
        validation.InputTitle = "Enter Value";
        validation.InputMessage = "Please enter a number between 0 and 500.";
        validation.ErrorTitle = "Invalid Input";
        validation.ErrorMessage = "The value must be between 0 and 500.";
        validation.ShowInput = true;
        validation.ShowError = true;
        validation.AlertStyle = ValidationAlertType.Stop;

        // Save the workbook with the validation applied
        workbook.Save("ColumnKValidation.xlsx");
    }
}
