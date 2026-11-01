// Title: Assign row‑specific multiplication formulas to column B while looping column A using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# loop through column and set formula per row | How to create a dynamic multiplication formula in Excel with Aspose.Cells .NET | Generate row‑based formulas in column B using the Aspose.Cells API | Calculate and save formulas after assigning them in an Aspose.Cells workbook | C# code to populate column A and apply =Arow*row formula with Aspose.Cells
// Tags: loop column assign formula Aspose.Cells | row based multiplication formula C# | programmatic cell formula setting .NET | calculate workbook formulas Aspose.Cells | save calculated Excel file .NET

using System;
using Aspose.Cells;

// The example creates a new workbook, fills cells A1‑A10 with values 1‑10, iterates through those rows to place a formula "=A{row}*{row}" in column B, calculates all formulas, and saves the result as MultiplicationResults.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate column A with sample values (1 to 10)
        for (int i = 0; i < 10; i++)
        {
            cells[i, 0].PutValue(i + 1); // A1..A10
        }

        // Loop through column A, assign incremental row references in a multiplication formula
        // Formula in column B: =A{row}*{row}
        for (int i = 0; i < 10; i++)
        {
            int rowIndex = i + 1; // Excel rows are 1‑based
            string formula = $"=A{rowIndex}*{rowIndex}";
            cells[i, 1].Formula = formula; // B column
        }

        // Calculate the formulas to store the results
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("MultiplicationResults.xlsx");
    }
}
