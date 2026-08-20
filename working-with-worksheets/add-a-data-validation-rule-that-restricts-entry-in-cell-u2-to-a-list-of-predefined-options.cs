// Title: C# – Add List‑Based Data Validation to Cell U2 with Aspose.Cells for .NET
// Description: Shows how to create a workbook, target cell U2, apply a List‑type validation containing predefined options, enable the in‑cell drop‑down, and save the file as U2_Validation.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# data validation | ValidationType.List | Excel drop‑down list | cell U2 | restrict cell input | predefined options | Aspose.Cells example | Excel automation | list validation
// Common Searches: Aspose.Cells add list validation C# | how to create drop‑down in a single Excel cell using Aspose | restrict cell U2 to specific values Aspose.Cells | C# code for ValidationType.List Aspose | Excel data validation with Aspose.Cells .NET
// Developer Intent: Apply a List validation to cell U2 so that only the specified options can be entered.
// Use Cases: Design a data‑entry template where users select a category from a drop‑down in U2. | Generate reports that enforce status values (e.g., Approved, Pending, Rejected) via list validation in U2. | Automate workbook creation for surveys, limiting answers in U2 to predefined choices. | Create a financial model that restricts currency codes in U2 to a fixed list.
// AI Prompts: Generate C# code with Aspose.Cells that adds a list validation to cell U2 containing 'OptionA', 'OptionB', 'OptionC' and saves as 'U2_Validation.xlsx'. | Show how to reference a named range on another worksheet for the validation list instead of a hard‑coded string. | Explain how to customize the error alert (title, message, style) for the list validation on cell U2. | Provide a step‑by‑step guide to enable the in‑cell drop‑down and set a custom input message for cell U2 validation.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    // Shows how to create a workbook, target cell U2, apply a List‑type validation containing predefined options, enable the in‑cell drop‑down, and save the file as U2_Validation.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the cell area for the validation (U2)
                CellArea area = new CellArea
                {
                    StartRow = 1,      // Row index is zero‑based (U2 -> row 2)
                    EndRow = 1,
                    StartColumn = 20, // Column index is zero‑based (U -> 21st column)
                    EndColumn = 20
                };

                // Add a new validation rule for the specified area
                int validationIndex = worksheet.Validations.Add(area);
                Validation validation = worksheet.Validations[validationIndex];

                // Set the validation type to a list of predefined options
                validation.Type = ValidationType.List;

                // Define the allowed values (comma‑separated) and enclose in double quotes
                validation.Formula1 = "\"OptionA,OptionB,OptionC\"";

                // Enable the in‑cell drop‑down list
                validation.InCellDropDown = true;

                // Save the workbook
                workbook.Save("U2_Validation.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
