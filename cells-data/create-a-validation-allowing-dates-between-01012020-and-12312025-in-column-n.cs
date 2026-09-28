// Title: Add date‑range validation (01/01/2020‑12/31/2025) to column N rows 1‑1000 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook and applies a 'Between' date validation to column N (rows 1‑1000) with lower bound 01/01/2020 and upper bound 12/31/2025. | Show how to configure custom input and error messages for a date validation rule on a specific worksheet column using Aspose.Cells in C#. | Generate a complete example that saves the workbook as an .xlsx file after adding the date validation to column N.
// Common Searches: Aspose.Cells C# set date validation for column N between 2020 and 2025 | How to restrict Excel column N to dates from Jan 1 2020 to Dec 31 2025 using Aspose.Cells | C# Aspose.Cells add data validation with between operator for a specific column range
// Tags: Aspose.Cells date validation between dates | C# column N validation Aspose.Cells | CellArea range validation .NET | Excel date range restriction Aspose.Cells | Custom input error messages Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsDateValidationExample
{
    // Creates a new workbook, defines a CellArea covering rows 1‑1000 of column N, adds a date validation of type 'Between' with bounds 01/01/2020 and 12/31/2025, sets user‑friendly input and error titles/messages, and saves the file as DateValidationColumnN.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range for column N (index 13) from row 1 to row 1000 (0‑based indices)
            CellArea dateValidationArea = CellArea.CreateCellArea(0, 13, 999, 13);

            // Add a data validation to the worksheet for the defined area
            int validationIndex = worksheet.Validations.Add(dateValidationArea);
            Validation dateValidation = worksheet.Validations[validationIndex];

            // Configure the validation to allow dates between 01/01/2020 and 12/31/2025
            dateValidation.Type = ValidationType.Date;                     // Date validation type
            dateValidation.Operator = OperatorType.Between;                // Between operator
            dateValidation.Formula1 = "01/01/2020";                         // Lower bound
            dateValidation.Formula2 = "12/31/2025";                         // Upper bound

            // Optional: user-friendly messages
            dateValidation.InputTitle = "Enter a Date";
            dateValidation.InputMessage = "Please enter a date between 01/01/2020 and 12/31/2025.";
            dateValidation.ErrorTitle = "Invalid Date";
            dateValidation.ErrorMessage = "The date must be within the allowed range.";
            dateValidation.ShowInput = true;
            dateValidation.ShowError = true;

            // Save the workbook to a file
            workbook.Save("DateValidationColumnN.xlsx");
        }
    }
}
