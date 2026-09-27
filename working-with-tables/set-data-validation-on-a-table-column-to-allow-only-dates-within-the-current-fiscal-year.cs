// Title: Apply fiscal‑year date validation to a ListObject column in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that builds a ListObject, determines the current fiscal year (July 1 – June 30), and enforces a date‑only entry rule on a specified column. | Adjust the validation so the target range expands automatically when new rows are added to the table. | Include custom input and error messages that show the exact start and end dates of the fiscal year.
// Common Searches: aspocells c# set date validation for a table column based on fiscal year | how to restrict Excel table column entries to dates within the current fiscal year using Aspose.Cells | dynamic data validation range for ListObject column in .NET workbook | c# Aspose.Cells date validation between two OADate values
// Tags: date validation fiscal year Aspose.Cells | ListObject column data validation .NET | dynamic validation range CellArea Aspose | OADate formula validation C# | Excel table column custom messages Aspose

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a new workbook, adds a ListObject named MyTable, computes the current fiscal year (July 1 – June 30), applies a date‑type validation to the DateColumn that only permits dates within that fiscal year, sets custom input and error messages, and saves the file as FiscalYearValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Define the table range (including header)
            int firstRow = 0;      // Row 1 (zero‑based)
            int firstCol = 0;      // Column A
            int lastRow = 9;       // Row 10
            int lastCol = 1;       // Column B

            // Add header values
            sheet.Cells[firstRow, firstCol].PutValue("DateColumn");
            sheet.Cells[firstRow, firstCol + 1].PutValue("ValueColumn");

            // Create a table (ListObject) over the defined range
            int tableIndex = sheet.ListObjects.Add(firstRow, firstCol, lastRow, lastCol, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Calculate the current fiscal year (July 1 – June 30)
            DateTime today = DateTime.Today;
            int fiscalYearStart = today.Month >= 7 ? today.Year : today.Year - 1;
            DateTime fiscalStart = new DateTime(fiscalYearStart, 7, 1);
            DateTime fiscalEnd = fiscalStart.AddYears(1).AddDays(-1); // June 30 of next year

            // Set data validation on the DateColumn (exclude header)
            int dataStartRow = firstRow + 1; // Row 2
            int dataEndRow = 99;             // Row 100 (adjust as needed)
            int dateColIndex = firstCol;     // Column A

            // Define the validation area using CellArea properties
            CellArea validationArea = new CellArea
            {
                StartRow = dataStartRow,
                StartColumn = dateColIndex,
                EndRow = dataEndRow,
                EndColumn = dateColIndex
            };

            int validationIndex = sheet.Validations.Add(validationArea);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.Date;
            validation.Operator = OperatorType.Between;
            // Aspose.Cells expects OADate values as strings for date formulas
            validation.Formula1 = fiscalStart.ToOADate().ToString(CultureInfo.InvariantCulture);
            validation.Formula2 = fiscalEnd.ToOADate().ToString(CultureInfo.InvariantCulture);
            validation.InputMessage = "Enter a date within the current fiscal year.";
            validation.ErrorMessage = $"Date must be between {fiscalStart:d} and {fiscalEnd:d}.";
            validation.ShowInput = true;
            validation.ShowError = true;

            // Save the workbook
            workbook.Save("FiscalYearValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
