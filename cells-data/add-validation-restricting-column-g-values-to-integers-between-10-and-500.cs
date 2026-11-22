// Title: Add whole-number validation (10‑500) to column G of an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to enforce integer values between 10 and 500 in column G | C# add data validation for a specific column in Excel with Aspose.Cells | Set whole number range validation on Excel column using Aspose.Cells .NET API | Validate Excel column G values are between 10 and 500 with Aspose.Cells | Apply numeric range restriction to a column in an existing workbook using Aspose.Cells
// Tags: Aspose.Cells whole number validation C# | Excel column G data validation Aspose.Cells | numeric range restriction Aspose.Cells .NET | add validation rule to worksheet Aspose.Cells | C# restrict Excel cell values between 10 and 500

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, defines the entire column G as a validation area, and adds a whole-number validation that permits only integers from 10 to 500. It configures error and input messages, ensures the output directory exists, saves the modified file, and includes basic file‑existence and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for column G (entire column)
            // Column G is the 7th column, index 6 (zero‑based)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = sheet.Cells.MaxDataRow, // apply to all used rows; you can also use 1048575 for full column
                StartColumn = 6,
                EndColumn = 6
            };

            // Add a validation rule to the worksheet
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;          // Restrict to whole numbers (integers)
            validation.Operator = OperatorType.Between;           // Value must be between two limits
            validation.Formula1 = "10";                            // Minimum allowed value
            validation.Formula2 = "500";                           // Maximum allowed value
            validation.ShowError = true;                           // Show error dialog when invalid data entered
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "Please enter an integer between 10 and 500 in column G.";
            // ShowInputMessage property may not be available in older versions; omitted for compatibility
            validation.InputTitle = "Input Required";
            validation.InputMessage = "Enter an integer between 10 and 500.";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
