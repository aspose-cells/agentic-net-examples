// Title: Aspose.Cells .NET: Add date‑range validation to column N (rows 1‑1000) for 2020‑2025
// Description: Creates a new workbook, defines a CellArea covering column N (index 13) rows 1‑1000, and adds a Validation of type Date with the Between operator. The rule limits entries to 01/01/2020 – 12/31/2025, includes custom input and error messages, and saves the file as DateValidationColumnN.xlsx.
// Keywords: Aspose.Cells | C# | date validation | Excel column N | date range 2020 2025 | Data validation API | CellArea | ValidationType.Date | OperatorType.Between | input message | error message
// Common Searches: Aspose.Cells set date validation for column N | C# add Excel date range validation 2020 to 2025 | How to restrict Excel column to specific dates using Aspose.Cells | Aspose.Cells data validation with custom messages | Create date validation for a column in .NET
// Developer Intent: Add a data‑validation rule that permits only dates between 01/01/2020 and 12/31/2025 in column N of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Ensure financial reports use only allowed fiscal periods by limiting column N to 2020‑2025 dates. | Prevent entry errors in a data‑entry template where column N stores transaction dates. | Provide users with clear guidance and error feedback when entering dates in Excel files generated programmatically.
// AI Prompts: Show how to replace the fixed dates with a dynamic range based on the current date in Aspose.Cells. | Generate C# code that applies the same 2020‑2025 date validation to columns N, O, and P. | Explain how to localize the input and error messages for multiple cultures in Aspose.Cells validation.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, defines a CellArea covering column N (index 13) rows 1‑1000, and adds a Validation of type Date with the Between operator. The rule limits entries to 01/01/2020 – 12/31/2025, includes custom input and error messages, and saves the file as DateValidationColumnN.xlsx.
    public class DateValidationInColumnN
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the range for column N (index 13) from row 1 to row 1000
                // CellArea uses zero‑based indexes: startRow = 0, startColumn = 13, endRow = 999, endColumn = 13
                CellArea validationArea = CellArea.CreateCellArea(0, 13, 999, 13);

                // Add a new validation to the worksheet for the defined area
                int validationIndex = worksheet.Validations.Add(validationArea);
                Validation validation = worksheet.Validations[validationIndex];

                // Configure the validation to allow only dates between 01/01/2020 and 12/31/2025
                validation.Type = ValidationType.Date;                     // Date validation
                validation.Operator = OperatorType.Between;               // Between operator
                validation.Formula1 = "01/01/2020";                       // Lower bound
                validation.Formula2 = "12/31/2025";                       // Upper bound

                // Optional: user messages
                validation.InputTitle = "Enter a date";
                validation.InputMessage = "Date must be between 01/01/2020 and 12/31/2025.";
                validation.ErrorTitle = "Invalid Date";
                validation.ErrorMessage = "The entered date is outside the allowed range.";
                validation.ShowInput = true;
                validation.ShowError = true;

                // Save the workbook
                workbook.Save("DateValidationColumnN.xlsx");
                Console.WriteLine("Workbook saved as DateValidationColumnN.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DateValidationInColumnN.Run();
        }
    }
}
