// Title: Add whole-number data validation to rows 2‑100 in column A using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a CellArea for rows 2 through 100 in column A and applies a WholeNumber validation with a Between operator ranging from 1 to 1000 using Aspose.Cells. | Show how to configure an Aspose.Cells Validation object to restrict values to whole numbers within a custom range for a specific column range in an Excel worksheet. | Generate a complete example that saves the workbook after adding numeric validation to the defined cell area.
// Common Searches: Aspose.Cells C# apply integer validation to a column range | Create CellArea for rows 2-100 column A and set numeric validation in Aspose.Cells | Configure between operator for integer validation in Aspose.Cells .NET | Example of adding data validation to specific rows using Aspose.Cells C#
// Tags: Aspose.Cells integer validation C# | CellArea range validation Aspose.Cells | Excel numeric validation between operator | C# Aspose.Cells data validation example

using System;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    // The sample creates a new workbook, defines a CellArea covering rows 2‑100 in column A, adds a whole-number validation with a Between operator limited to 1‑1000, and saves the file as WholeNumberValidation.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a CellArea that covers rows 2‑100 (zero‑based indices 1‑99) in column A (column index 0)
            CellArea validationArea = CellArea.CreateCellArea(1, 0, 99, 0);

            // Add a validation to the worksheet for the defined area
            int validationIndex = sheet.Validations.Add(validationArea);
            Validation validation = sheet.Validations[validationIndex];

            // Configure the validation as Whole Number
            validation.Type = ValidationType.WholeNumber;
            // Optional: restrict the whole numbers to a range, e.g., 1‑1000
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "1000";

            // Save the workbook
            workbook.Save("WholeNumberValidation.xlsx", SaveFormat.Xlsx);
        }
    }
}
