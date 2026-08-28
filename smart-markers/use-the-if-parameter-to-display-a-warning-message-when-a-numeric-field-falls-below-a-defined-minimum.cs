// Title: Add a warning‑style numeric minimum validation with custom error and input messages to a cell range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, fills column A, and applies a whole‑number validation to A1:A10 that shows a warning when the entered value is less than 10 using Aspose.Cells. | Generate a C# snippet to configure ValidationAlertType.Warning with a custom error title and message for numeric cells that fall below a defined threshold in an Aspose.Cells worksheet. | Provide a C# example that adds an input prompt and warning message to a numeric validation rule for a column of cells in an Excel file using Aspose.Cells.
// Common Searches: asp.net add warning data validation for numbers below minimum using Aspose.Cells | c# Aspose.Cells set numeric validation with custom error title and message | how to show input message with numeric validation in Aspose.Cells workbook | Aspose.Cells less than operator warning alert example in C#
// Tags: Aspose.Cells warning style validation | C# set validation alert type warning | custom error title Aspose.Cells | input prompt for validated cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, populates column A with numeric values, adds a whole‑number validation to cells A1:A10 that triggers a warning when a value is less than 10, and configures custom error title, error message, and an input prompt before saving the file as NumericMinimumWarningDemo.xlsx.
    public class NumericMinimumWarningDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate some sample numeric data in column A (cells A1:A10)
                for (int i = 0; i < 10; i++)
                {
                    worksheet.Cells[i, 0].PutValue(i * 5); // values: 0,5,10,...45
                }

                // Define the range to which the validation will be applied (A1:A10)
                CellArea validationArea = CellArea.CreateCellArea(0, 0, 9, 0);

                // Add a validation rule to the worksheet using the newer API
                int validationIndex = worksheet.Validations.Add(validationArea);
                Validation validation = worksheet.Validations[validationIndex];

                // Configure the validation:
                // - Whole number type
                // - Operator: LessThan (value must be less than the minimum)
                // - Minimum threshold set to 10
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.LessThan;
                validation.Formula1 = "10";

                // Set the alert style to Warning so a warning message is shown (not a stop)
                validation.AlertStyle = ValidationAlertType.Warning;

                // Provide a title and message for the warning
                validation.ErrorTitle = "Value Too Low";
                validation.ErrorMessage = "The entered number is below the allowed minimum of 10.";

                // Ensure the warning is displayed when the user enters an invalid value
                validation.ShowError = true;

                // Optionally, display an input message when the cell is selected
                validation.ShowInput = true;
                validation.InputTitle = "Enter Value";
                validation.InputMessage = "Please enter a number greater than or equal to 10.";

                // Save the workbook
                workbook.Save("NumericMinimumWarningDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            NumericMinimumWarningDemo.Run();
        }
    }
}
