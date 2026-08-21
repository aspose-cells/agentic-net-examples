// Title: C# – Add Whole‑Number (1‑100) Data Validation to Column C with Aspose.Cells
// Description: Creates a new workbook, defines a CellArea for column C, adds a validation of type WholeNumber with Operator Between, sets the lower and upper bounds to 1 and 100, configures optional input and error messages, and saves the file as ColumnCWholeNumberValidation.xlsx.
// Keywords: Aspose.Cells | C# data validation | Excel whole number validation | column C validation | range 1-100 | CellArea | ValidationType.WholeNumber | OperatorType.Between | input message | error message | .NET example
// Common Searches: Aspose.Cells set data validation column C | C# restrict Excel cell to whole numbers 1‑100 | add input and error messages Aspose.Cells | Excel validation example Aspose.Cells .NET | how to apply whole‑number validation with Aspose.Cells
// Developer Intent: Add a validation rule that permits only whole numbers between 1 and 100 in column C of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Ensure quantity fields in generated reports contain only valid whole numbers. | Prevent downstream processing errors by blocking out‑of‑range entries in a template. | Guide end‑users with clear input prompts and error alerts while filling a data‑entry sheet.
// AI Prompts: Generate C# code that applies a whole‑number (1‑100) validation to column C with custom input and error messages using Aspose.Cells. | Show how to define a dynamic CellArea for column C and attach a Between validation in Aspose.Cells. | Explain how to modify the validation range, bounds, or messages after the workbook has been created.

using Aspose.Cells;

// Creates a new workbook, defines a CellArea for column C, adds a validation of type WholeNumber with Operator Between, sets the lower and upper bounds to 1 and 100, configures optional input and error messages, and saves the file as ColumnCWholeNumberValidation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the validation area for column C (zero‑based column index 2)
        // Here we apply the rule to rows 0 through 1000; adjust as needed
        CellArea area = CellArea.CreateCellArea(0, 2, 1000, 2);

        // Add a new validation to the worksheet for the defined area
        int validationIndex = worksheet.Validations.Add(area);
        Validation validation = worksheet.Validations[validationIndex];

        // Configure the validation: whole numbers between 1 and 100
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "1";
        validation.Formula2 = "100";

        // Optional user messages
        validation.InputTitle = "Whole Number Required";
        validation.InputMessage = "Please enter a whole number between 1 and 100.";
        validation.ErrorTitle = "Invalid Input";
        validation.ErrorMessage = "The value must be a whole number between 1 and 100.";
        validation.ShowInput = true;
        validation.ShowError = true;

        // Save the workbook with the validation rule applied
        workbook.Save("ColumnCWholeNumberValidation.xlsx");
    }
}
