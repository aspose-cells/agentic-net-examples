// Title: Create email address validation with input prompt and stop‑style error alert for cell A1 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a custom validation to cell A1 which checks that the value contains an '@' before a '.' and shows an input message prompting for a valid email address. | Show how to configure Aspose.Cells to display a stop‑style error alert with a custom title and message when the email validation formula fails. | Provide a complete example that creates a workbook, applies the email validation with both input and error messages, and saves the file as EmailValidationDemo.xlsx.
// Common Searches: Aspose.Cells C# add custom email validation with input message to a specific cell | C# set data validation for email format using Aspose.Cells workbook | Show input prompt for cell validation in Aspose.Cells .NET example | Configure stop alert for invalid email entry in Aspose.Cells worksheet | Apply custom formula validation for email address in Aspose.Cells C#
// Tags: email format custom validation Aspose.Cells | display input prompt for validation C# | validation error alert stop style Aspose.Cells | specific cell validation setup Aspose.Cells | FIND function email pattern C#

using System;
using Aspose.Cells;

// The example creates a new workbook, defines cell A1, adds a custom validation formula that ensures an '@' appears before a '.' in the entered text, displays an input prompt asking for a valid email address, shows a stop‑style error alert with custom titles when the pattern is not met, and saves the workbook as EmailValidationDemo.xlsx.
public class EmailValidationDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell area (A1) where the validation will be applied
            CellArea area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 0, EndColumn = 0 };

            // Add a validation to the defined area
            int validationIndex = worksheet.Validations.Add(area);
            Validation validation = worksheet.Validations[validationIndex];

            // Set validation type to Custom with a simple email pattern check
            // The formula ensures the cell contains "@" and "." with "@" appearing before "."
            validation.Type = ValidationType.Custom;
            // SetFormula1 requires isR1C1 and isArrayFormula flags; both are false for standard A1 notation
            validation.SetFormula1("=AND(ISNUMBER(FIND(\"@\",A1)),ISNUMBER(FIND(\".\",A1)),FIND(\"@\",A1)<FIND(\".\",A1)))", false, false);

            // Show input message when the cell is selected
            validation.ShowInput = true;
            validation.InputTitle = "Email Input";
            validation.InputMessage = "Please enter a valid email address (e.g., user@example.com)";

            // Show error alert if the entered value does not satisfy the formula
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Email";
            validation.ErrorMessage = "The value entered is not a valid email address.";
            validation.AlertStyle = ValidationAlertType.Stop;

            // Save the workbook to a file
            workbook.Save("EmailValidationDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        EmailValidationDemo.Run();
    }
}
