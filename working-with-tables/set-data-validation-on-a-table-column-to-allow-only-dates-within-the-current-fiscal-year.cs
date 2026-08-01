// Title: C# Aspose.Cells – Add fiscal‑year date validation to a table column
// Description: The sample builds a new workbook, defines a CellArea for column B (rows 2‑100), adds a Validation object, sets it to Date with a Between operator, calculates the current fiscal year (July 1 – June 30), assigns the start and end dates as OADate strings, configures input and error messages, and saves the worksheet.
// Keywords: Aspose.Cells | C# date validation | Excel fiscal year validation | data validation between dates | set validation for table column | OADate | current fiscal year | Excel worksheet validation .NET
// Common Searches: Aspose.Cells set date validation for fiscal year | C# restrict Excel column to fiscal year dates | How to add date range validation in Aspose.Cells | Excel data validation July to June using Aspose.Cells | Apply data validation to a table column in .NET
// Developer Intent: Create a validation rule that permits only dates falling within the active fiscal year for a designated worksheet column.
// Use Cases: Ensure expense entries in a financial report belong to the current fiscal year. | Limit project start dates in a schedule to the active fiscal year. | Prevent out‑of‑range dates in a budgeting table column. | Validate timesheet dates so they align with the organization’s fiscal calendar.
// AI Prompts: Generate C# code with Aspose.Cells that adds a date validation for column C rows 5‑200, restricting entries to the current fiscal year. | Explain how to modify the example to use a custom fiscal‑year start month instead of July. | Provide a snippet that reads fiscal start and end dates from a JSON configuration file and applies them to an Aspose.Cells validation rule.

using System;
using Aspose.Cells;

namespace FiscalYearDateValidationApp
{
    // The sample builds a new workbook, defines a CellArea for column B (rows 2‑100), adds a Validation object, sets it to Date with a Between operator, calculates the current fiscal year (July 1 – June 30), assigns the start and end dates as OADate strings, configures input and error messages, and saves the worksheet.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define the data‑validation range (column B, rows 2‑100)
                CellArea dateColumnArea = new CellArea
                {
                    StartRow = 1,      // Row 2 (zero‑based)
                    EndRow = 99,       // Row 100
                    StartColumn = 1,   // Column B
                    EndColumn = 1
                };

                // Add a validation object for the defined area
                int validationIndex = sheet.Validations.Add(dateColumnArea);
                Validation dateValidation = sheet.Validations[validationIndex];

                // Set validation to Date type with Between operator
                dateValidation.Type = ValidationType.Date;
                dateValidation.Operator = OperatorType.Between;

                // Determine the current fiscal year (July 1 – June 30)
                DateTime today = DateTime.Today;
                DateTime fiscalStart;
                DateTime fiscalEnd;

                if (today.Month >= 7) // July or later
                {
                    fiscalStart = new DateTime(today.Year, 7, 1);
                    fiscalEnd   = new DateTime(today.Year + 1, 6, 30);
                }
                else // before July
                {
                    fiscalStart = new DateTime(today.Year - 1, 7, 1);
                    fiscalEnd   = new DateTime(today.Year, 6, 30);
                }

                // Assign lower and upper bounds as OADate literals (Excel numeric dates)
                dateValidation.Formula1 = fiscalStart.ToOADate()
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);
                dateValidation.Formula2 = fiscalEnd.ToOADate()
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);

                // Optional user messages
                dateValidation.InputTitle = "Fiscal Year Date";
                dateValidation.InputMessage = $"Enter a date between {fiscalStart:yyyy-MM-dd} and {fiscalEnd:yyyy-MM-dd}.";
                dateValidation.ErrorTitle = "Invalid Date";
                dateValidation.ErrorMessage = "The date is outside the current fiscal year.";
                dateValidation.ShowInput = true;
                dateValidation.ShowError = true;

                // Save the workbook
                string outputPath = "FiscalYearDateValidation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
