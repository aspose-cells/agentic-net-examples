// Title: Add a calculated Category column with a nested IF formula to classify sales as High, Medium, or Low using Aspose.Cells for C#
// Common Searches: how to add a calculated column with nested IF in Aspose.Cells C# | Aspose.Cells C# categorize sales values high medium low | set IF formula for each row in an Excel worksheet using Aspose.Cells | programmatically create a Category column in an Excel workbook with Aspose.Cells | apply conditional classification to sales data using Aspose.Cells C#
// Tags: add calculated column Aspose.Cells C# | nested IF formula Excel Aspose.Cells | sales category classification Excel C# | set cell formula CellsHelper Aspose | generate high medium low labels Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an existing workbook, inserts a new "Category" column, assigns a nested IF formula to classify each sales value from column B as "High", "Medium", or "Low", recalculates the formulas, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("SalesData.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Assume sales values are in column B (index 1) starting from row 2 (index 1)
        // We'll add a new column "Category" in column C (index 2)
        int salesColumnIndex = 1;      // B column
        int categoryColumnIndex = 2;   // C column
        int headerRowIndex = 0;        // Row 1 (zero‑based)

        // Set header for the new column
        sheet.Cells[headerRowIndex, categoryColumnIndex].PutValue("Category");

        // Determine the last used row in the sheet
        int lastRow = sheet.Cells.MaxDataRow;

        // Loop through each data row and set the IF formula
        for (int row = 1; row <= lastRow; row++)
        {
            // Build the IF formula:
            // =IF(B2>1000,"High",IF(B2>500,"Medium","Low"))
            string salesCellRef = CellsHelper.CellIndexToName(salesColumnIndex, row);
            string formula = $"IF({salesCellRef}>1000,\"High\",IF({salesCellRef}>500,\"Medium\",\"Low\"))";

            // Assign the formula to the Category cell
            sheet.Cells[row, categoryColumnIndex].Formula = formula;
        }

        // Recalculate all formulas to store the results (optional, as Aspose.Cells does it on save)
        workbook.CalculateFormula();

        // Save the updated workbook
        workbook.Save("SalesData_WithCategory.xlsx");
    }
}
