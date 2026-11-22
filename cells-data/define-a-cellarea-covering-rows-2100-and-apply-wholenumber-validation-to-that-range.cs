// Title: Create a CellArea for rows 2‑100 and add whole‑number data validation with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set whole number validation for rows 2 through 100 | How to define a CellArea covering rows 2 to 100 in Aspose.Cells | Apply numeric range validation to a column range using Aspose.Cells .NET | Add data validation to multiple rows in an Excel workbook with Aspose.Cells | C# Aspose.Cells create validation for rows 2‑100 across all columns
// Tags: CellArea rows 2-100 Aspose.Cells | integer validation with Aspose.Cells .NET | numeric range validation between operator | apply data validation to multiple rows Aspose.Cells | Aspose.Cells example adding validation formulas

using System;
using Aspose.Cells;

// Creates a new workbook, defines a CellArea covering rows 2‑100 across all columns, adds an integer (whole‑number) validation that restricts values between 1 and 1000, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a CellArea covering rows 2‑100 (zero‑based indices 1‑99) for all columns
            CellArea area = new CellArea
            {
                StartRow = 1,      // Row 2
                EndRow = 99,       // Row 100
                StartColumn = 0,   // Column A
                EndColumn = 255    // Up to column IV (covers typical columns)
            };

            // Add a whole‑number validation to the defined area
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";      // Minimum allowed value
            validation.Formula2 = "1000";   // Maximum allowed value
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "Please enter a whole number between 1 and 1000.";

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
