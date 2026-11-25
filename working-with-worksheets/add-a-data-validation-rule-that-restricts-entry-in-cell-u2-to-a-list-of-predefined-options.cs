// Title: Create an in‑cell dropdown list validation for cell U2 with predefined options using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to restrict cell U2 to a predefined list of values | Add data validation list to a specific cell in an Excel workbook using Aspose.Cells .NET | C# generate dropdown list validation in Excel cell with Aspose.Cells | Set in‑cell dropdown for cell U2 with custom options using Aspose.Cells for .NET | Configure validation error message for list‑type validation in Aspose.Cells C#
// Tags: Aspose.Cells list validation .NET | Excel cell dropdown validation C# | predefined options validation Aspose.Cells | in‑cell dropdown list configuration | validation error alert customization

using System;
using Aspose.Cells;

// The example creates a new workbook, obtains the Validation object for cell U2, sets its type to List, defines a comma‑separated list of allowed values, enables the in‑cell dropdown, adds optional input prompts and error alerts, and saves the workbook as Validation_U2.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (empty)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Target cell: U2
            const string targetCell = "U2";

            // Obtain the Validation object for the target cell
            Validation validation = worksheet.Cells[targetCell].GetValidation();

            // Set validation type to List
            validation.Type = ValidationType.List;

            // Define the allowed options (comma‑separated, enclosed in double quotes)
            validation.Formula1 = "\"Option1,Option2,Option3\"";

            // Enable the in‑cell dropdown list
            validation.InCellDropDown = true;

            // Optional: provide an input prompt
            validation.ShowInput = true;
            validation.InputTitle = "Select an option";
            validation.InputMessage = "Choose one of the predefined values.";

            // Optional: configure an error alert
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid entry";
            validation.ErrorMessage = "The value must be one of the listed options.";

            // Save the workbook to disk
            const string outputPath = "Validation_U2.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
