using System;
using Aspose.Cells;

class RetrieveFormula
{
    static void Main()
    {
        // Path to the Excel file containing the target cell
        string filePath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet (adjust index or name if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the cell at address E10
        Cell cell = worksheet.Cells["E10"];

        // Retrieve the formula in standard A1 notation (non‑localized)
        string formula = cell.GetFormula(false, false);

        // Log the retrieved formula for audit purposes
        Console.WriteLine($"Formula in cell E10: {formula}");
    }
}