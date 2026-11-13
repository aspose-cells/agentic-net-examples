// Title: Add a calculated “Category” column with IF formulas to classify sales in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to insert a calculated column with IF statement in Aspose.Cells C# | Aspose.Cells C# categorize sales values into high, medium, low using formulas | add IF formula to existing Excel sheet with Aspose.Cells .NET | recalculate all formulas after adding a new column in Aspose.Cells | save workbook with new calculated field using Aspose.Cells for .NET
// Tags: add calculated column IF formula Aspose.Cells | sales value categorization Excel Aspose.Cells | recalculate formulas workbook Aspose.Cells | save workbook with new column Aspose.Cells | insert IF statement into cells Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook, creates a "Category" header, and writes an IF formula for each row that classifies the sales amount in column B as High (>1000), Medium (>500), or Low. After inserting the formulas, it recalculates the workbook and saves the result as a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("SalesData.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Assume sales figures are in column B (zero‑based index 1) starting from row 2 (index 1)
        int firstDataRow = 1; // Row 2 in Excel
        int lastDataRow = sheet.Cells.MaxDataRow; // Last row that contains data

        // Add header for the calculated field in column D (zero‑based index 3)
        sheet.Cells[0, 3].PutValue("Category");

        // Insert IF formula for each data row to categorize sales:
        // > 1000 => "High", > 500 => "Medium", otherwise "Low"
        for (int row = firstDataRow; row <= lastDataRow; row++)
        {
            // Build the formula string using Excel A1 notation
            string formula = $"IF(B{row + 1}>1000,\"High\",IF(B{row + 1}>500,\"Medium\",\"Low\"))";
            sheet.Cells[row, 3].Formula = formula;
        }

        // Recalculate all formulas so the new column contains actual values
        workbook.CalculateFormula();

        // Save the updated workbook
        workbook.Save("SalesData_WithCategory.xlsx");
    }
}
