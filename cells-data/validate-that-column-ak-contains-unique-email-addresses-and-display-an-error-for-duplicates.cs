// Title: Enforce unique email addresses in Excel column AK with Aspose.Cells for .NET using custom data validation
// Common Searches: Aspose.Cells C# add custom validation to prevent duplicate emails in column AK | How to set a uniqueness rule for an email column in an Excel file using Aspose.Cells .NET | C# Aspose.Cells COUNTIF formula to ensure each email in column AK is unique | Validate duplicate email addresses in Excel with Aspose.Cells and display an error alert | Create data validation that blocks duplicate emails in column AK using Aspose.Cells
// Tags: custom validation unique email column Aspose.Cells | COUNTIF uniqueness rule Excel .NET | error alert duplicate email Aspose.Cells | column AK data validation Aspose.Cells | prevent duplicate emails Excel C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, defines a CellArea covering column AK, adds a custom validation rule with a COUNTIF formula that forces each email to appear only once, configures a stop‑style error alert titled "Duplicate Email", and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or specify the required one)
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the last used row in the sheet (zero‑based index)
                int lastRow = sheet.Cells.MaxDataRow;

                // Define the area covering column AK (index 36) from the first row to the last used row
                CellArea emailArea = new CellArea
                {
                    StartRow = 0,
                    EndRow = lastRow,
                    StartColumn = 36,   // AK = 37th column (zero‑based index 36)
                    EndColumn = 36
                };

                // Add a custom data‑validation rule to the defined area
                int validationIndex = sheet.Validations.Add(emailArea);
                Validation emailValidation = sheet.Validations[validationIndex];

                // Set the validation type to Custom and provide a formula that ensures uniqueness
                emailValidation.Type = ValidationType.Custom;
                // Formula checks that the current email appears only once in the whole column
                emailValidation.Formula1 = $"COUNTIF($AK$1:$AK${lastRow + 1},AK1)=1";

                // Configure the error alert that will be shown when a duplicate is entered
                emailValidation.ShowError = true;
                // Default alert style is Stop, which blocks entry; explicit setting omitted for compatibility
                emailValidation.ErrorTitle = "Duplicate Email";
                emailValidation.ErrorMessage = "The email address must be unique in column AK.";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to the desired output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
