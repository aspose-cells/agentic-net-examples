// Title: Create date validation for column N to accept only dates from 01/01/2020 through 12/31/2025 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set date validation for entire column N between 2020 and 2025 | How to restrict Excel column N to a specific date range with Aspose.Cells .NET | C# add between operator date validation to column N using Aspose.Cells | Define CellArea for column N and apply date range validation in Aspose.Cells | Excel data validation for dates 01/01/2020-12/31/2025 with Aspose.Cells C#
// Tags: Aspose.Cells date validation between specific dates | C# column N CellArea validation Aspose.Cells | Excel date range validation OperatorType.Between | Aspose.Cells workbook save with validation | Data validation for Excel column N using .NET

using Aspose.Cells;
using System;

// The example creates a new workbook, defines the full column N with a CellArea, adds a date validation of type Date using the Between operator, sets the allowed range from 01/01/2020 to 12/31/2025, configures an error message, and saves the file as DateValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for column N (N1:N1048576)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 1048575,
                StartColumn = 13, // Column N (0‑based index)
                EndColumn = 13
            };

            // Add date validation to the defined range
            Validation validation = sheet.Validations[sheet.Validations.Add(area)];
            validation.Type = ValidationType.Date;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1/1/2020";
            validation.Formula2 = "12/31/2025";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Date";
            validation.ErrorMessage = "Please enter a date between 01/01/2020 and 12/31/2025.";

            // Save the workbook
            workbook.Save("DateValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
