using System;
using Aspose.Cells;

class ErrorCheckOptionsDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the collection of error‑check options for the worksheet
        ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;

        // Add a new ErrorCheckOption to the collection
        int optionIdx = errorCheckOptions.Add();
        ErrorCheckOption errorCheckOption = errorCheckOptions[optionIdx];

        // Configure which error types should be checked (true = check, false = ignore)
        errorCheckOption.SetErrorCheck(ErrorCheckType.NumberStoredAsText, false); // ignore numbers stored as text
        errorCheckOption.SetErrorCheck(ErrorCheckType.InconsistFormula, false);   // ignore inconsistent formulas
        errorCheckOption.SetErrorCheck(ErrorCheckType.TextDate, true);           // enable text‑date check

        // Define a cell area (A1:C5) and associate it with this error‑check option
        CellArea area = CellArea.CreateCellArea("A1", "C5");
        errorCheckOption.AddRange(area);

        // Display the number of ranges attached to this option (should be 1)
        Console.WriteLine("Number of ranges in the error‑check option: " + errorCheckOption.GetCountOfRange());

        // Save the workbook to a file
        workbook.Save("ErrorCheckOptionsDemo.xlsx");
    }
}