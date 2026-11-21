// Title: Create row‑specific multiplication formulas in column B with Aspose.Cells for .NET
// Common Searches: how to assign a formula to each cell in a column using Aspose.Cells C# | generate incremental multiplication formulas based on row index with Aspose.Cells | loop through column A and set =Arow*row formula in Aspose.Cells .NET | calculate all formulas after adding them programmatically in Aspose.Cells | find the last used row in a worksheet using Aspose.Cells C#
// Tags: dynamic cell formula assignment Aspose.Cells | row‑based multiplication formula C# | retrieve last data row Aspose.Cells | calculate workbook formulas Aspose.Cells | save workbook as XLSX Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, fills column A with sequential numbers, determines the last occupied row, and then loops through each row to place a formula in column B that multiplies the A‑column value by the row's 1‑based index (e.g., =A1*1, =A2*2). After assigning all formulas it triggers calculation and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate column A with sample data (1, 2, 3, ...)
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1);
        }

        // Determine the last used row in column A
        int lastRow = sheet.Cells.MaxDataRow;

        // Loop through column A and assign an incremental multiplication formula in column B
        for (int row = 0; row <= lastRow; row++)
        {
            // Formula uses the current row number (1‑based) as the multiplier
            // Example for row 0 (Excel row 1): =A1*1
            string formula = $"=A{row + 1}*{row + 1}";
            sheet.Cells[row, 1].Formula = formula;
        }

        // Calculate all formulas so that column B contains the results
        workbook.CalculateFormula();

        // Save the workbook (save rule)
        workbook.Save("output.xlsx");
    }
}
