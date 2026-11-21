// Title: Add a whole-number (1‑100) data validation rule to column C in an Excel worksheet with Aspose.Cells for .NET
// Common Searches: how to apply whole number validation to column C using Aspose.Cells in C# | Aspose.Cells set data validation range for entire column in .NET | C# create between operator validation for Excel cells with Aspose.Cells | configure custom error and input messages for Aspose.Cells validation | determine last used row and apply validation in Aspose.Cells workbook
// Tags: Aspose.Cells numeric range validation | Excel column C validation Aspose.Cells | ValidationType WholeNumber operator Between example | CellArea column definition Aspose.Cells | custom validation messages Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, calculates the last populated row, defines a CellArea covering column C, adds a WholeNumber validation with a Between operator limited to 1‑100, sets custom error and input prompts, and saves the workbook as DataValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last row that contains data (0‑based). If the sheet is empty, use row 0.
            int lastRow = sheet.Cells.MaxDataRow;
            if (lastRow < 0) lastRow = 0;

            // Define the range for the entire column C (zero‑based column index 2)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = lastRow,
                StartColumn = 2,
                EndColumn = 2
            };

            // Add a data‑validation rule for whole numbers between 1 and 100
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;   // restrict to whole numbers
            validation.Operator = OperatorType.Between;    // value must be between two limits
            validation.Formula1 = "1";                     // lower bound
            validation.Formula2 = "100";                   // upper bound

            // Optional UI settings
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "Please enter a whole number between 1 and 100.";
            // ShowInputMessage property is not available in this version; setting input message is sufficient
            validation.InputTitle = "Enter Number";
            validation.InputMessage = "Enter a whole number from 1 to 100.";

            // Save the workbook
            string outputPath = "DataValidation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
