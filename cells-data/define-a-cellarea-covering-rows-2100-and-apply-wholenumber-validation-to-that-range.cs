// Title: Add Whole‑Number Validation to Rows 2‑100 in Column A with Aspose.Cells for .NET
// Description: Demonstrates how to create a CellArea covering rows 2‑100 in column A, attach a whole‑number validation rule (between 1 and 1000), and save the workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells CellArea rows 2-100 | C# whole number validation Aspose.Cells | CreateCellArea example .NET | Validation OperatorType.Between | Save workbook after validation
// Common Searches: Aspose.Cells add whole number validation to a range | C# define CellArea rows 2 to 100 column A | Set validation operator between values Aspose.Cells | Save Excel file after applying validation Aspose
// Developer Intent: Define a CellArea that spans rows 2‑100 in column A and apply a whole‑number validation rule to that range using Aspose.Cells for .NET.
// Use Cases: Ensure employee ID entries in column A (rows 2‑100) are whole numbers within a permitted range. | Validate quantity fields in an inventory sheet so only integer values between 1 and 1000 are accepted. | Prevent non‑numeric data entry in financial reports by restricting a column block to whole numbers.
// AI Prompts: Generate C# code with Aspose.Cells to create a CellArea for rows 2‑100 in column B and apply whole‑number validation between 10 and 500. | Explain how to modify the validation to allow only positive integers for a dynamic row range using Aspose.Cells. | Show how to retrieve an existing validation rule from a worksheet and update its operator or formula values programmatically.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    // Demonstrates how to create a CellArea covering rows 2‑100 in column A, attach a whole‑number validation rule (between 1 and 1000), and save the workbook using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a cell area that covers rows 2‑100 in column A (zero‑based indices)
            // StartRow = 1 (row 2), EndRow = 99 (row 100), StartColumn = EndColumn = 0 (column A)
            CellArea validationArea = CellArea.CreateCellArea(1, 0, 99, 0);

            // Add a whole‑number validation to the defined area
            int validationIndex = sheet.Validations.Add(validationArea);
            Validation validation = sheet.Validations[validationIndex];

            // Configure the validation as a whole‑number type
            validation.Type = ValidationType.WholeNumber;
            // Optional: restrict the whole numbers to a specific range, e.g., 1‑1000
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "1000";

            // Save the workbook
            workbook.Save("WholeNumberValidation.xlsx", SaveFormat.Xlsx);
        }
    }
}
