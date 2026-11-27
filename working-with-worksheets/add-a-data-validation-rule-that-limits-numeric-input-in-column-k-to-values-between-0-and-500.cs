// Title: Create a whole-number data validation (0‑500) for the entire column K in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add whole number validation to column K in Excel | How to restrict numeric input between 0 and 500 for an entire column using Aspose.Cells .NET | Set data validation for whole numbers on column K with Aspose.Cells workbook | C# Aspose.Cells create validation rule for Excel column K range | Apply Between operator validation for column K using Aspose.Cells API
// Tags: Aspose.Cells whole-number validation API | Excel column K validation range Aspose.Cells | OperatorType Between data validation | CellArea column validation Aspose.Cells | numeric range 0-500 validation C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, defines a CellArea that covers the entire column K (rows 1‑1,048,576), adds a whole-number validation restricting values to 0‑500, configures an input prompt and error alert, and saves the file as DataValidation_ColumnK.xlsx.
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

                // Define the range for column K (entire column)
                // Using the maximum rows for Excel (1,048,576)
                CellArea validationRange = new CellArea
                {
                    StartRow = 0,
                    EndRow = 1048575,
                    StartColumn = 10, // Column K (0‑based index)
                    EndColumn = 10
                };

                // Add a data validation rule for whole numbers between 0 and 500
                // Add returns the index of the created validation; retrieve the object afterwards
                int validationIndex = sheet.Validations.Add(validationRange);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "0";
                validation.Formula2 = "500";

                // Optional: set input message and error alert
                validation.InputMessage = "Enter a number between 0 and 500.";
                validation.ErrorMessage = "Invalid entry! Value must be between 0 and 500.";
                validation.ShowError = true; // Show error alert

                // Prepare output path and ensure directory exists
                string outputPath = "DataValidation_ColumnK.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
