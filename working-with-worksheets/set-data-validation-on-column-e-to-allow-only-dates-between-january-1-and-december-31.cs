// Title: Apply date‑range data validation to the entire column E in an Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that adds a Validation of type Date to column E (E1:E1048576) and restricts entries to dates between 1/1/2023 and 12/31/2023. | Generate a complete example that sets a Between operator, custom input message, and error alert for date validation on column E, then saves the workbook.
// Common Searches: Aspose.Cells C# how to restrict column E to a specific date range in Excel | set data validation for whole column E to allow only dates in 2023 using Aspose.Cells | C# Aspose.Cells date validation between Jan 1 and Dec 31 for column E | apply input and error messages for date validation on column E with Aspose.Cells .NET
// Tags: Aspose.Cells date validation column | C# set validation for Excel column E | Aspose.Cells ValidationType.Date example | Excel date range validation using Aspose.Cells | Aspose.Cells between operator for date validation

using Aspose.Cells;
using System;
using System.IO;

// Creates a new workbook, defines the full range of column E, adds a Date‑type Validation with the Between operator limited to 1/1/2023‑12/31/2023, sets custom input and error messages, and saves the file as DataValidationExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for column E (E1:E1048576)
            // Use CellArea.CreateCellArea to build the range object
            CellArea area = CellArea.CreateCellArea("E1", "E1048576");

            // Add a data validation rule for the specified range
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];

            // Configure validation to allow dates between Jan 1 and Dec 31, 2023
            validation.Type = ValidationType.Date;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1/1/2023";
            validation.Formula2 = "12/31/2023";

            // Optional: display messages to the user
            validation.InputMessage = "Enter a date between Jan 1 and Dec 31, 2023.";
            validation.ShowError = true;
            validation.ErrorMessage = "Invalid date. Please enter a date within the allowed range.";

            // Prepare output path
            string outputPath = "DataValidationExample.xlsx";
            string fullPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullPath);

            // Create directory if needed (outputDir can be null when only a file name is provided)
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {fullPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
