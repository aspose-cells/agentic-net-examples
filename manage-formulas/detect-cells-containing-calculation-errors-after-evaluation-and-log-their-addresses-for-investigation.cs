using System;
using Aspose.Cells;

class DetectErrorCells
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Enable calculation chain to allow dependent queries (optional but safe)
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Calculate all formulas without ignoring errors so they surface
        workbook.CalculateFormula(new CalculationOptions() { IgnoreError = false });

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the used range of the worksheet
        int maxRow = sheet.Cells.MaxDataRow;
        int maxCol = sheet.Cells.MaxDataColumn;

        // Iterate through each cell in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = sheet.Cells[row, col];

                // Check if the cell's calculated result is an error
                if (cell.IsErrorValue)
                {
                    // Retrieve the cell address (e.g., "A1")
                    string address = cell.Name;

                    // Obtain the specific error type, if needed
                    ErrorCellValueType errorType = cell.GetRichValue().ErrorValue;

                    // Log the error information
                    Console.WriteLine($"Error detected in cell {address}: {errorType}");
                }
            }
        }

        // Save the workbook after processing (optional)
        workbook.Save("output.xlsx");
    }
}