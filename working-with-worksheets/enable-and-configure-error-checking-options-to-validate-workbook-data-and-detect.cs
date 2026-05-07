using System;
using Aspose.Cells;

namespace AsposeCellsErrorCheckDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data that may trigger error checks
            sheet.Cells["A1"].PutValue("123");                     // Number stored as text
            sheet.Cells["B1"].Formula = "=SUM(A1:A10)";            // May cause evaluation error if range empty
            sheet.Cells["C1"].PutValue("01/01/20");                // Text date
            sheet.Cells["D1"].PutValue("TextNumber");             // Text that looks like a number

            // Get the collection of error‑check options for the worksheet
            ErrorCheckOptionCollection errorOptions = sheet.ErrorCheckOptions;

            // Add a new error‑check option
            int optionIndex = errorOptions.Add();
            ErrorCheckOption option = errorOptions[optionIndex];

            // Configure which error types should be checked (true = show green triangle on error)
            option.SetErrorCheck(ErrorCheckType.NumberStoredAsText, true);   // Enable check for numbers stored as text
            option.SetErrorCheck(ErrorCheckType.EvaluationError, true);     // Enable check for formula evaluation errors
            option.SetErrorCheck(ErrorCheckType.TextDate, true);            // Enable check for text dates
            option.SetErrorCheck(ErrorCheckType.TextNumber, true);          // Enable check for text numbers
            option.SetErrorCheck(ErrorCheckType.InconsistFormula, false);  // Disable inconsistent formula check

            // Apply the error‑check option to the whole used range of the worksheet
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;
            CellArea fullRange = CellArea.CreateCellArea(0, 0, maxRow, maxCol);
            option.AddRange(fullRange);

            // Save the workbook (the error‑check settings are stored in the file)
            workbook.Save("ErrorCheckConfigured.xlsx");
        }
    }
}