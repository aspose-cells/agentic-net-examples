// Title: Add a list‑type data validation dropdown to cell U2 in a new workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a List validation to cell U2, defines options OptionA, OptionB, OptionC, sets a custom error title and message, and saves the workbook as Output.xlsx. | Write a C# snippet using Aspose.Cells to apply a dropdown list validation to cell U2, include an input prompt, configure error handling, and export the file.
// Common Searches: Aspose.Cells C# how to create a dropdown list validation for a single cell | Set list validation options in cell U2 with Aspose.Cells for .NET | C# Aspose.Cells validation.Formula1 syntax for comma‑separated list values | Customize error title and message for data validation in Aspose.Cells workbook | Add input message to list validation in Aspose.Cells C# example
// Tags: Aspose.Cells list‑type data validation C# | specific cell validation Aspose.Cells | validation.Formula1 comma‑separated list Aspose.Cells | validation error caption Aspose.Cells | input message configuration Aspose.Cells validation

using System;
using Aspose.Cells;

// The program creates a new workbook, applies a list‑type data validation dropdown to cell U2 with predefined options (OptionA, OptionB, OptionC), configures custom error and input messages, and saves the file as Output.xlsx.
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

            // Define the cell area for U2 (row 2, column U)
            CellArea area = CellArea.CreateCellArea("U2", "U2");

            // Add a data validation rule to the specified cell
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];

            // Set validation type to List and define the allowed options
            validation.Type = ValidationType.List;
            // The list of options must be enclosed in double quotes and separated by commas
            validation.Formula1 = "\"OptionA,OptionB,OptionC\"";

            // Configure error messages
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Entry";
            validation.ErrorMessage = "Please select a value from the predefined list.";

            // Configure input (prompt) message – only InputMessage is available in this API version
            validation.InputMessage = "Choose one of the allowed options.";

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
