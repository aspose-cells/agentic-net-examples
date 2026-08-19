// Title: Aspose.Cells C# – Retrieve validation type and formulas for cell J5
// Description: C# example that adds a whole‑number validation (10‑20) to J5, uses Worksheet.Cells["J5"].GetValidation() to read the Validation.Type, Formula1 and Formula2, logs them, and saves the workbook.
// Keywords: Aspose.Cells C# | GetValidation | cell validation | validation type | Formula1 | Formula2 | J5 | Aspose.Cells .NET example | GitHub source code | Aspose.Cells API
// Common Searches: Aspose.Cells get validation of a cell C# | How to read validation type with Aspose.Cells | Retrieve cell J5 validation formulas Aspose.Cells | Worksheet.Cells GetValidation example | C# Aspose.Cells validation Type and formulas
// Developer Intent: Read and display the validation settings applied to a specific cell (J5) in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Verify that required data‑validation rules exist before distributing a spreadsheet. | Dynamically adjust UI components based on a cell's validation constraints. | Export validation metadata to audit logs or reporting tools.
// AI Prompts: Write C# code that loops through a range and prints each cell's validation type, Formula1 and Formula2 using Aspose.Cells. | Show how to modify a Validation object retrieved from a cell (e.g., change operator or formulas) and reapply it. | Explain how to handle null results from GetValidation when a cell has no validation applied.

using System;
using Aspose.Cells;

// C# example that adds a whole‑number validation (10‑20) to J5, uses Worksheet.Cells["J5"].GetValidation() to read the Validation.Type, Formula1 and Formula2, logs them, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell area (J5) where the validation will be applied
            // J = column index 9 (0‑based), 5 = row index 4 (0‑based)
            CellArea validationArea = new CellArea
            {
                StartRow = 4,
                EndRow = 4,
                StartColumn = 9,
                EndColumn = 9
            };

            // Add a validation for the specified area
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "10";
            validation.Formula2 = "20";

            // Retrieve the validation applied to cell J5
            Validation retrievedValidation = worksheet.Cells["J5"].GetValidation();

            // Log validation details
            Console.WriteLine("Validation Type: " + retrievedValidation.Type);
            Console.WriteLine("Formula1: " + retrievedValidation.Formula1);
            Console.WriteLine("Formula2: " + retrievedValidation.Formula2);

            // Save the workbook
            workbook.Save("ValidationJ5.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
