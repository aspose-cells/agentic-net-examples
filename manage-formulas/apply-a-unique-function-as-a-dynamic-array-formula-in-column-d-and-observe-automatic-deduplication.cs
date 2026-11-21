// Title: Use Aspose.Cells for .NET to apply the UNIQUE dynamic‑array formula and automatically deduplicate rows in Excel
// Common Searches: how to apply the UNIQUE function in a C# Aspose.Cells workbook | deduplicate rows with UNIQUE dynamic array formula using Aspose.Cells .NET | force calculation of dynamic array formulas like UNIQUE in Aspose.Cells | spilling results of the UNIQUE formula to adjacent columns with Aspose.Cells | set formula =UNIQUE(A1:C6) programmatically in Aspose.Cells C#
// Tags: UNIQUE dynamic‑array formula Aspose.Cells | deduplicate Excel rows C# Aspose.Cells | calculate formulas programmatically Aspose.Cells | spilling array results workbook .NET | set cell formula Aspose.Cells API

using Aspose.Cells;
using System;

// The example creates a new workbook, fills A1:C6 with sample data, assigns the formula =UNIQUE(A1:C6) to D1 to automatically deduplicate rows, forces formula calculation, and saves the result as UniqueDemo.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data placed in columns A, B, C (rows 1‑6)
        string[,] sampleData = new string[,]
        {
            { "Apple",  "Red",    "Fruit" },
            { "Banana", "Yellow", "Fruit" },
            { "Apple",  "Red",    "Fruit" },
            { "Carrot", "Orange", "Vegetable" },
            { "Banana", "Yellow", "Fruit" },
            { "Broccoli","Green","Vegetable" }
        };

        // Fill the worksheet with the sample data
        for (int row = 0; row < sampleData.GetLength(0); row++)
        {
            for (int col = 0; col < sampleData.GetLength(1); col++)
            {
                sheet.Cells[row, col].PutValue(sampleData[row, col]);
            }
        }

        // Apply the UNIQUE dynamic‑array formula in column D starting at D1.
        // The formula deduplicates the rows from A1:C6 and spills the result into D onward.
        sheet.Cells["D1"].Formula = "=UNIQUE(A1:C6)";

        // Force calculation so the UNIQUE function is evaluated.
        workbook.CalculateFormula();

        // Save the workbook to verify the deduplicated result.
        workbook.Save("UniqueDemo.xlsx");
    }
}
