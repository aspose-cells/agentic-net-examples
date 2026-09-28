// Title: Add numeric range validation (0‑500) to the whole column K with Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that applies a whole‑number validation to every cell in column K, limiting entries to values from 0 to 500 and showing custom input and error messages using Aspose.Cells. | Write a C# example that defines a CellArea for column K, adds a between‑operator validation rule, configures the messages, and saves the workbook as an .xlsx file.
// Common Searches: aspocells c# whole number validation for column K 0 to 500 | apply data validation to an entire Excel column using Aspose.Cells .NET | c# aspocells set custom input and error messages for numeric validation
// Tags: Aspose.Cells whole number validation column K | C# numeric limits data validation Excel | Aspose.Cells between operator validation | Excel input and error messages with Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, defines a CellArea that spans all rows of column K, adds a whole‑number validation with a between operator restricting values to 0‑500, sets custom input and error messages, and saves the file as OutputWithValidation.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define the cell area for column K (zero‑based index 10) across all rows
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = sheet.Cells.MaxRow,   // 1,048,575
                    StartColumn = 10,              // Column K
                    EndColumn = 10
                };

                // Add a validation rule for the defined area
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];

                // Set validation type to whole number and define limits
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "0";   // Lower bound
                validation.Formula2 = "500"; // Upper bound

                // Input and error messages
                validation.InputMessage = "Please enter a number between 0 and 500.";
                validation.ErrorMessage = "The value must be between 0 and 500.";
                validation.ShowInput = true;   // Show the input message when the cell is selected
                validation.ShowError = true;   // Show the error message when invalid data is entered

                // Save the workbook
                string outputPath = "OutputWithValidation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
