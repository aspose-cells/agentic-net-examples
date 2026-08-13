// Title: C# – Add fiscal‑year date validation to an Excel table column with Aspose.Cells
// Description: Creates a workbook, defines a ListObject named "DateTable", calculates the start and end dates of the current fiscal year (Jan 1‑Dec 31), and applies a Date validation (Operator: Between) to the table's data column. Includes optional input and error messages, then saves the file as FiscalYearDateValidation.xlsx.
// Keywords: Aspose.Cells date validation C# | Excel table column validation | fiscal year date restriction | ListObject data validation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set date validation for table column | C# restrict Excel column to current fiscal year dates | Add date range validation to ListObject using Aspose.Cells | Excel fiscal year validation Aspose.Cells .NET
// Developer Intent: Apply a data‑validation rule to a table column so that only dates falling within the current fiscal year are allowed.
// Use Cases: Enforce fiscal‑year dates in financial reporting templates. | Prevent out‑of‑range dates in project schedule tables. | Provide a ready‑to‑use Excel template with built‑in date constraints for data entry.
// AI Prompts: Generate C# code with Aspose.Cells that adds a date validation to a ListObject column limited to the current fiscal year. | Show how to change the fiscal year start month from January to a custom month in the validation logic. | Demonstrate applying the same fiscal‑year date validation to multiple columns of the same table.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, defines a ListObject named "DateTable", calculates the start and end dates of the current fiscal year (Jan 1‑Dec 31), and applies a Date validation (Operator: Between) to the table's data column. Includes optional input and error messages, then saves the file as FiscalYearDateValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a header and some sample dates (rows 2‑10)
            sheet.Cells["A1"].PutValue("Date");
            for (int i = 2; i <= 10; i++)
            {
                // Sample dates around today
                sheet.Cells[i - 1, 0].PutValue(DateTime.Now.AddDays(i - 5));
            }

            // Convert the range A1:A10 into a table (ListObject)
            // Use overload with hasHeaders parameter (true because A1 is a header)
            int tableIdx = sheet.ListObjects.Add(0, 0, 10, 1, true);
            ListObject table = sheet.ListObjects[tableIdx];
            // Set the display name of the table
            table.DisplayName = "DateTable";

            // Determine the start and end dates of the current fiscal year (Jan 1 – Dec 31)
            DateTime now = DateTime.Now;
            int fiscalYear = now.Year;
            DateTime fiscalStart = new DateTime(fiscalYear, 1, 1);
            DateTime fiscalEnd   = new DateTime(fiscalYear, 12, 31);

            // Define the validation area: the data column of the table (exclude header)
            CellArea dateColumnArea = new CellArea
            {
                StartRow = 1, // Row 2 in Excel (zero‑based index)
                EndRow   = 9, // Row 10
                StartColumn = 0,
                EndColumn   = 0
            };

            // Add a validation rule to the worksheet for the defined area
            int validationIdx = sheet.Validations.Add(dateColumnArea);
            Validation validation = sheet.Validations[validationIdx];

            // Configure the validation to allow only dates within the fiscal year
            validation.Type = ValidationType.Date;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = fiscalStart.ToString("yyyy-MM-dd"); // lower bound
            validation.Formula2 = fiscalEnd.ToString("yyyy-MM-dd");   // upper bound

            // Optional UI messages
            validation.ShowInput = true;
            validation.InputTitle = "Fiscal Year Date";
            validation.InputMessage = $"Enter a date between {fiscalStart:yyyy-MM-dd} and {fiscalEnd:yyyy-MM-dd}.";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Date";
            validation.ErrorMessage = "The date is outside the current fiscal year.";

            // Save the workbook
            workbook.Save("FiscalYearDateValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
