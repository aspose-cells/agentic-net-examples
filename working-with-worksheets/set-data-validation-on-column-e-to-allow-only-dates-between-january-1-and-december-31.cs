// Title: C# – Add Date‑Range Validation to Column E in an Aspose.Cells Workbook
// Description: Demonstrates how to create a new Workbook with Aspose.Cells for .NET, define a CellArea covering column E, add a Validation object, set its type to Date with a Between operator, specify start and end dates (e.g., 1/1/2023 to 12/31/2023), configure optional input/error messages, and save the file as DateValidationColumnE.xlsx.
// Keywords: Aspose.Cells C# date validation | Excel column E data validation .NET | set date range validation Aspose.Cells | ValidationType.Date example C# | apply data validation to a column Aspose | Aspose.Cells workbook validation tutorial
// Common Searches: Aspose.Cells how to restrict column E to dates | C# add date range validation in Excel with Aspose | set data validation for specific column using Aspose.Cells | date between validation Aspose.Cells .NET | example of ValidationType.Date in C#
// Developer Intent: Create a validation rule that permits only dates between January 1 and December 31 in column E of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Ensure users enter only dates within the current year in a data‑entry column. | Prevent out‑of‑range dates when importing external records into a template. | Programmatically apply the same yearly date restriction across multiple sheets.
// AI Prompts: Generate C# code with Aspose.Cells to enforce a 2024 date range in columns F‑H. | Show how to duplicate an existing Validation object and assign it to another column. | Provide an example of custom input and error messages for a date validation rule in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a new Workbook with Aspose.Cells for .NET, define a CellArea covering column E, add a Validation object, set its type to Date with a Between operator, specify start and end dates (e.g., 1/1/2023 to 12/31/2023), configure optional input/error messages, and save the file as DateValidationColumnE.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the validation area for column E (zero‑based column index 4)
        // Here we apply it to rows 0 through 1000; adjust as needed
        CellArea validationArea = CellArea.CreateCellArea(0, 4, 1000, 4);

        // Add the validation to the worksheet
        int validationIndex = worksheet.Validations.Add(validationArea);
        Validation validation = worksheet.Validations[validationIndex];

        // Configure the validation to allow only dates between Jan 1 and Dec 31
        validation.Type = ValidationType.Date;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "1/1/2023";   // start date
        validation.Formula2 = "12/31/2023"; // end date

        // Optional user messages
        validation.InputTitle = "Date Entry";
        validation.InputMessage = "Please enter a date between Jan 1 and Dec 31, 2023.";
        validation.ErrorTitle = "Invalid Date";
        validation.ErrorMessage = "The date must be within the year 2023.";
        validation.ShowInput = true;
        validation.ShowError = true;

        // Save the workbook
        workbook.Save("DateValidationColumnE.xlsx");
    }
}
