// Title: Restrict Column G to Whole Numbers 10‑500 with Aspose.Cells for .NET
// Description: Shows how to create a workbook, define a CellArea for column G (rows 1‑1000), add a WholeNumber validation using the Between operator, set the allowed range to 10‑500, configure custom prompts, and save the file.
// Keywords: Aspose.Cells | .NET | C# Excel validation | column G data validation | whole number between 10 and 500 | Excel data validation API | CellArea | ValidationType.WholeNumber | OperatorType.Between
// Common Searches: Aspose.Cells set data validation C# | C# restrict Excel column values 10-500 | how to add whole number validation in Aspose.Cells | Excel column G validation Aspose | Aspose.Cells validation example .NET | data validation for column G using Aspose.Cells | Aspose.Cells tutorial integer range validation
// Developer Intent: Add a data‑validation rule that permits only integer values from 10 to 500 in column G of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Prevent out‑of‑range amounts in financial spreadsheets, ensuring entries stay within the 10‑500 limit. | Create a template for inventory imports where counts in column G must fall between 10 and 500, reducing manual correction. | Generate standardized Excel forms with built‑in validation to guide users and lower entry errors across departments.
// AI Prompts: Write C# code with Aspose.Cells that applies a WholeNumber validation (10‑500) to column G for rows 1‑1000 and includes custom input and error messages. | Explain how to modify the validation to reference cells A1 and A2 for dynamic minimum and maximum values instead of hard‑coded numbers. | Provide a script that copies the same integer‑range validation to every worksheet in a workbook and saves the result.

using Aspose.Cells;

// Shows how to create a workbook, define a CellArea for column G (rows 1‑1000), add a WholeNumber validation using the Between operator, set the allowed range to 10‑500, configure custom prompts, and save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the validation area for column G (index 6), rows 0 to 999
        CellArea area = CellArea.CreateCellArea(0, 6, 999, 6);

        // Add a validation to the worksheet for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Set validation to allow whole numbers between 10 and 500
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "10";
        validation.Formula2 = "500";

        // Optional user messages
        validation.InputTitle = "Enter Integer";
        validation.InputMessage = "Please enter an integer between 10 and 500.";
        validation.ErrorTitle = "Invalid Input";
        validation.ErrorMessage = "Value must be an integer between 10 and 500.";
        validation.ShowInput = true;
        validation.ShowError = true;

        // Save the workbook
        workbook.Save("ColumnGValidation.xlsx");
    }
}
