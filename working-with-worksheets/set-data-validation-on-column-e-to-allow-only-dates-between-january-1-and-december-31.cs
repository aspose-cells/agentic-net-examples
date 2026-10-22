// Title: Add date‑range data validation to the whole column E in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, applies a date validation rule to column E restricting entries to dates between January 1 and December 31 of a specified year, and saves the file. | Adjust the existing validation to use a different year or a custom date format while keeping the rule applied to the entire column E in Aspose.Cells. | Add custom input and error messages for the date validation on column E and show how to enable or disable the in‑cell dropdown in Aspose.Cells C#.
// Common Searches: Aspose.Cells C# set date validation for an entire column in Excel | How to restrict column E to dates between Jan 1 and Dec 31 using Aspose.Cells .NET | C# Aspose.Cells add data validation with between operator for dates | Change date range validation year in Aspose.Cells workbook programmatically | Add custom input and error messages to Excel date validation with Aspose.Cells C#
// Tags: Aspose.Cells date validation column | C# Excel data validation between dates | apply validation to whole column Aspose.Cells | custom input error messages Aspose.Cells | set validation operator between Aspose.Cells | date range validation .NET Excel

using Aspose.Cells;
using System;

// Creates a new workbook, defines a date‑type validation for the entire column E that only permits dates from 1 Jan to 31 Dec (2023), configures input and error messages, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for the entire column E (E1:E1048576).
            CellArea area = new CellArea
            {
                StartRow = 0,          // Row 1 (zero‑based)
                EndRow = 1048575,      // Last possible row in Excel
                StartColumn = 4,       // Column E (zero‑based)
                EndColumn = 4
            };

            // Add a data validation rule to the defined range.
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.Date;               // Allow only dates
            validation.Operator = OperatorType.Between;          // Between two dates
            validation.Formula1 = "1/1/2023";                     // Start date
            validation.Formula2 = "12/31/2023";                  // End date
            validation.IgnoreBlank = true;
            validation.InCellDropDown = false;

            // Show input (prompt) message.
            // Note: ShowInputMessage property may not be available in older Aspose.Cells versions.
            // The message will appear when the user selects a cell in the range.
            validation.InputMessage = "Enter a date between Jan 1 and Dec 31.";

            // Show error message when validation fails.
            validation.ShowError = true;
            validation.ErrorMessage = "Invalid date. Please enter a date within the allowed range.";

            // Save the workbook.
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
