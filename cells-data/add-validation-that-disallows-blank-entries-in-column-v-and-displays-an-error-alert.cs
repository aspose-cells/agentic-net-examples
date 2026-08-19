// Title: C# – Aspose.Cells: Prevent Blank Entries in Column V with Data Validation and Error Alert
// Description: Shows how to use Aspose.Cells for .NET to add an AnyValue validation to column V (index 21), reject blank cells, and display a Stop‑style error message when a user attempts to leave the cell empty. The workbook is saved as ColumnVValidation.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel data validation | column V | prevent blank cells | error alert | ValidationType.AnyValue | IgnoreBlank false | Stop alert | programmatic workbook
// Common Searches: Aspose.Cells column V validation C# | How to block blank cells in Excel using Aspose.Cells | Set custom error message for data validation Aspose.Cells .NET | Create required field validation in Excel with C# | Stop style validation Aspose.Cells
// Developer Intent: Create a data‑validation rule for column V that disallows blank values and shows a custom error alert in an Excel file generated with Aspose.Cells for .NET.
// Use Cases: Generating a template where column V must contain a product code. | Ensuring mandatory fields are filled in automated financial reports. | Applying required‑field validation across multiple sheets in a bulk‑export process. | Providing end‑users with immediate feedback when they leave a required cell empty.
// AI Prompts: Write C# code using Aspose.Cells to add a required‑field validation to column V with a custom Stop‑style error title and message. | Show how to determine the last used row and set the validation range dynamically. | Explain how to switch the alert style from Stop to Warning or Information and update the error text. | Provide a version of the code that applies the same validation rule to multiple worksheets in a workbook using Aspose.Cells. | Generate a PowerShell script that invokes the compiled C# example to create the validated workbook.

using Aspose.Cells;

// Shows how to use Aspose.Cells for .NET to add an AnyValue validation to column V (index 21), reject blank cells, and display a Stop‑style error message when a user attempts to leave the cell empty. The workbook is saved as ColumnVValidation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the validation area for column V (zero‑based index 21)
        // Here we apply it to rows 0 through 1000; adjust as needed
        CellArea area = CellArea.CreateCellArea(0, 21, 1000, 21);

        // Add a new validation to the worksheet for the defined area
        int validationIndex = worksheet.Validations.Add(area);
        Validation validation = worksheet.Validations[validationIndex];

        // Use AnyValue type and do NOT ignore blanks, thus blank entries are invalid
        validation.Type = ValidationType.AnyValue;
        validation.IgnoreBlank = false;

        // Configure the error alert that will be shown when a blank is entered
        validation.AlertStyle = ValidationAlertType.Stop;   // Stop style prevents entry
        validation.ErrorTitle = "Blank Not Allowed";
        validation.ErrorMessage = "Please enter a value in column V.";
        validation.ShowError = true;                       // Show the error message

        // Save the workbook with the validation applied
        workbook.Save("ColumnVValidation.xlsx");
    }
}
