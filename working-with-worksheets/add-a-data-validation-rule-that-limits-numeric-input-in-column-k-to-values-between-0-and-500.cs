// Title: Add whole-number data validation (0‑500) to the entire column K in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set whole number validation for column K between 0 and 500 | How to restrict values in an entire Excel column using Aspose.Cells .NET | C# add numeric data validation to column K with Aspose.Cells | Apply between operator data validation to a column in Excel via Aspose.Cells | Example of adding data validation to a full column in Aspose.Cells for .NET
// Tags: Aspose.Cells whole number validation | C# Excel column data validation | Aspose.Cells validation operator between | Excel numeric range restriction .NET | Apply validation to entire column Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace DataValidationExample
{
    // The sample creates a new workbook, defines the full range of column K, adds a whole‑number validation rule that only permits values from 0 to 500, sets an error message, and saves the worksheet as DataValidationExample.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with a default worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define the cell area for the entire column K (0‑based column index 10)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 1048575, // Maximum rows in Excel 2007+
                    StartColumn = 10,
                    EndColumn = 10
                };

                // Add a validation for the defined area
                ValidationCollection validations = sheet.Validations;
                int validationIndex = validations.Add(area); // use the non‑obsolete overload
                Validation validation = validations[validationIndex];

                // Configure validation: whole number between 0 and 500
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "0";
                validation.Formula2 = "500";

                // Optional error message displayed to the user
                validation.ShowError = true;
                validation.ErrorTitle = "Invalid Input";
                validation.ErrorMessage = "Please enter a number between 0 and 500.";

                // Save the workbook
                string outputPath = "DataValidationExample.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
