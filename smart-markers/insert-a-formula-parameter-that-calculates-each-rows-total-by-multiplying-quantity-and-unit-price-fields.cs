// Title: Insert a per‑row Total column with a quantity × unit‑price formula using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing .xlsx file with Aspose.Cells, iterates over the data range, and writes a formula =A(row)*B(row) into column C for every row. | Demonstrate how to programmatically create a Total column by assigning a multiplication expression to cells using the Aspose.Cells Cells API. | Provide a self‑contained example that persists the workbook after embedding the per‑row total formulas.
// Common Searches: Aspose.Cells C# add formula to each row in existing worksheet | Calculate total column (quantity times price) with Aspose.Cells .NET | How to set cell formula while iterating rows using Aspose.Cells | Insert dynamic multiplication formula into Excel using Aspose.Cells for .NET | Save workbook after adding calculated column with Aspose.Cells C#
// Tags: Aspose.Cells set cell formula C# | Aspose.Cells insert calculated column Excel | Aspose.Cells loop rows assign formula | Aspose.Cells save modified workbook .NET | Aspose.Cells multiply columns A B total

using System;
using Aspose.Cells;

// Loads an existing Excel file, iterates over each data row, inserts a formula in column C that multiplies the values from columns A and B to compute a total, and saves the workbook as a new file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the used range to know how many rows contain data
        Cells cells = sheet.Cells;
        int firstDataRow = 1; // Assuming row 0 has headers, data starts at row 1 (zero‑based index)
        int lastDataRow = cells.MaxDataRow; // Zero‑based index of the last row with data

        // Insert formula in the "Total" column (e.g., column C, index 2)
        for (int row = firstDataRow; row <= lastDataRow; row++)
        {
            // Quantity in column A (index 0), Unit Price in column B (index 1)
            // Formula: =A{row+1}*B{row+1} (Excel rows are 1‑based)
            string formula = $"=A{row + 1}*B{row + 1}";
            cells[row, 2].Formula = formula;
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
