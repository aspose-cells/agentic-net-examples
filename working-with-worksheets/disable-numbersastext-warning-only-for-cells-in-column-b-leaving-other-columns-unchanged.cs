using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class DisableNumbersAsTextWarning
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // lifecycle: create
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the collection of error‑check options for the worksheet
            ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;

            // Add a new error‑check option entry
            int optionIndex = errorCheckOptions.Add();
            ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];

            // Disable the "Numbers stored as text" warning (TextNumber) for this option
            errorCheckOption.SetErrorCheck(ErrorCheckType.TextNumber, false);

            // Define a range that covers the entire column B (from row 1 to the maximum row)
            CellArea columnBRange = CellArea.CreateCellArea("B1", "B1048576");

            // Apply the option only to the defined range (column B)
            errorCheckOption.AddRange(columnBRange);

            // Save the workbook (lifecycle: save)
            workbook.Save("NumbersAsTextDisabled_ColumnB.xlsx");
        }
    }
}