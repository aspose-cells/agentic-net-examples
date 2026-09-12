// Title: Create a cumulative running‑total column in an Excel worksheet with Aspose.Cells for .NET using row‑referencing formulas
// AI Prompts: Generate C# code that uses Aspose.Cells to insert a formula in column B that adds the current row's value in column A to the previous row's running total. | Show how to programmatically assign a running‑total formula to each cell in a column, evaluate all formulas, and save the workbook with Aspose.Cells. | Demonstrate building a sample table of amounts and automatically creating a cumulative sum column using Aspose.Cells' Formula property in C#.
// Common Searches: Aspose.Cells C# set running total formula that references previous row | How to calculate cumulative sum in Excel using Aspose.Cells .NET API | Create dynamic running total column in Excel workbook with Aspose.Cells code example | Aspose.Cells formula for cumulative total across rows in C# | Programmatically add cumulative sum column to Excel sheet using Aspose.Cells
// Tags: aspocells cumulative sum formula | aspocells set cell formula previous row | aspocells calculate running total | excel running total column aspocells | c# aspocells evaluate formulas

using Aspose.Cells;
using System;

// Creates a new workbook, fills column A with sample amounts, writes a running‑total formula in column B that references the previous row, calculates all formulas, and saves the file as RunningTotal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Header row
        cells["A1"].PutValue("Amount");
        cells["B1"].PutValue("Running Total");

        // Sample data in column A (rows 2‑6)
        double[] amounts = { 100, 250, 150, 300, 200 };
        for (int i = 0; i < amounts.Length; i++)
        {
            // Row index in Aspose.Cells is zero‑based
            cells[i + 1, 0].PutValue(amounts[i]); // Column A
        }

        // Set running‑total formulas in column B
        int firstDataRow = 2; // Excel row number where data starts
        for (int excelRow = firstDataRow; excelRow < firstDataRow + amounts.Length; excelRow++)
        {
            if (excelRow == firstDataRow)
            {
                // First total equals the first amount: B2 = A2
                cells[excelRow - 1, 1].Formula = $"=A{excelRow}";
            }
            else
            {
                // Subsequent totals add previous total: B3 = B2 + A3, etc.
                cells[excelRow - 1, 1].Formula = $"=B{excelRow - 1}+A{excelRow}";
            }
        }

        // Evaluate all formulas
        workbook.CalculateFormula();

        // Save the workbook (lifecycle rule)
        workbook.Save("RunningTotal.xlsx");
    }
}
