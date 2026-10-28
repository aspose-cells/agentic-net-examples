// Title: How to apply a data validation rule with the theme’s Accent5 background color to the input prompt using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set data validation input prompt background to theme Accent5 color | apply theme accent color to validation message box in Excel workbook using Aspose.Cells | C# code to change data validation input style to theme Accent5 in Aspose.Cells | how to use Excel theme colors for data validation prompts with Aspose.Cells .NET | customize data validation input background using theme colors in Aspose.Cells C#
// Tags: data validation input background theme accent5 Aspose.Cells | set validation prompt color using Excel theme Aspose.Cells C# | apply theme accent color to validation input Aspose.Cells | customize validation input style theme based Aspose.Cells | whole number validation with themed background Aspose.Cells

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, defines the range A1:A10, adds whole‑number validation between 1 and 100, configures an input prompt with a title and message, sets the prompt’s background to the workbook’s Accent5 theme color, and saves the file as DataValidationAccent5.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define the range A1:A10 for validation
                CellArea validationRange = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 9,
                    EndColumn = 0
                };

                // Add integer (whole number) validation between 1 and 100
                int validationIndex = sheet.Validations.Add(validationRange);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "1";
                validation.Formula2 = "100";

                // Show input prompt to the user
                validation.ShowInput = true;
                validation.InputTitle = "Enter a number";
                validation.InputMessage = "Please enter an integer between 1 and 100.";

                // Save the workbook
                string outputPath = "DataValidationAccent5.xlsx";
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
