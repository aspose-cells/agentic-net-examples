// Title: Create a CellArea for rows 2‑100 and enforce whole‑number validation across all columns with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to set whole number validation for rows 2 through 100 | Define a CellArea covering rows 2 to 100 and apply integer data validation in .NET | Add data validation to an entire column range using Aspose.Cells validation API | Set input and error messages for whole number validation in an Excel workbook with Aspose.Cells | Apply Between operator integer validation to a large cell range in C# Aspose.Cells
// Tags: Aspose.Cells integer validation C# | CellArea rows 2-100 definition | Excel whole number data validation Aspose.Cells | Between operator validation Aspose.Cells | Custom input error messages validation Aspose

using Aspose.Cells;

// The example creates a new workbook, defines a CellArea that spans rows 2‑100 and all columns, adds a whole‑number validation using the Between operator with the full 32‑bit integer range, sets optional input and error messages, and saves the file as WholeNumberValidation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define a CellArea that covers rows 2‑100 (zero‑based index 1‑99)
        // and all columns from A to the maximum column (0‑255)
        CellArea area = new CellArea
        {
            StartRow = 1,      // Row 2
            EndRow = 99,       // Row 100
            StartColumn = 0,   // Column A
            EndColumn = 255    // Up to column IV (max column in Excel)
        };

        // Add a validation rule for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Set the validation type to WholeNumber
        validation.Type = ValidationType.WholeNumber;

        // Use the Between operator with a very wide range to allow any whole number
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "-2147483648"; // Minimum 32‑bit integer
        validation.Formula2 = "2147483647";  // Maximum 32‑bit integer

        // Optional: set input and error messages
        validation.InputMessage = "Please enter a whole number.";
        validation.ErrorMessage = "The value must be a whole number.";

        // Save the workbook to a file
        workbook.Save("WholeNumberValidation.xlsx");
    }
}
