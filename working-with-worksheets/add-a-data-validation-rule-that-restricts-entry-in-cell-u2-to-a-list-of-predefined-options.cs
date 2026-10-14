// Title: Create a list‑type data validation drop‑down in cell U2 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a list validation containing "Apple, Banana, Cherry" to cell U2 and saves the workbook. | Show how to configure a single‑cell data validation with an in‑cell drop‑down, input title, and error alert for column U row 2 using Aspose.Cells. | Generate a method that sets ValidationType.List for a specific CellArea (U2) from a string array and persists the Excel file.
// Common Searches: Aspose.Cells C# how to add a drop‑down list validation to a single cell | set data validation list for cell U2 using Aspose.Cells .NET | C# example of restricting Excel cell input to predefined options with Aspose.Cells | create in‑cell drop‑down for specific cell in Aspose.Cells workbook | Aspose.Cells validation formula for list of values in C#
// Tags: Aspose.Cells ValidationType.List C# | single‑cell validation rule Aspose.Cells | Excel list selection via Aspose.Cells | cell U2 validation configuration Aspose.Cells | validation formula from string array Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, defines an array of allowed values (Apple, Banana, Cherry), and applies a list‑type data validation with an in‑cell drop‑down, custom input title, and error alert to cell U2. The workbook is then saved as DataValidationExample.xlsx.
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

            // Define the list of allowed values
            string[] allowedValues = new string[] { "Apple", "Banana", "Cherry" };

            // Target cell U2 (zero‑based indices)
            int firstRow = 1;      // row 2
            int firstColumn = 20;  // column U

            // Define the cell area for the validation
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                StartColumn = firstColumn,
                EndRow = firstRow,
                EndColumn = firstColumn
            };

            // Add a new validation rule for the specified area (Add returns the index)
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];

            // Configure validation as a list
            validation.Type = ValidationType.List;
            validation.InCellDropDown = true;
            validation.Formula1 = "\"" + string.Join(",", allowedValues) + "\"";

            // Optional input message
            validation.InputTitle = "Select a value";
            validation.InputMessage = "Please choose from the list.";

            // Optional error alert
            validation.ErrorTitle = "Invalid entry";
            validation.ErrorMessage = "The value you entered is not in the allowed list.";

            // Save the workbook
            string outputPath = "DataValidationExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
