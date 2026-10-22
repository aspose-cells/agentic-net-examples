// Title: How to add a cumulative running‑total column with a SUM formula loop using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to insert a running total formula in each row of an Excel sheet | loop through rows and set SUM($B$2:Brow) formula with Aspose.Cells .NET | create financial worksheet with cumulative totals using Aspose.Cells C# | automate running total column in Excel workbook via Aspose.Cells API | calculate running totals per row in Aspose.Cells workbook programmatically
// Tags: Aspose.Cells cumulative sum loop | C# assign cell formula dynamically | Excel workbook SUM formula per row | financial sheet total calculation Aspose.Cells | programmatic formula insertion .NET

using System;
using Aspose.Cells;

// The C# program creates a new workbook, adds a "Financial" worksheet with Date, Amount, and Running Total headers, populates rows 2‑10 with dates and sample amounts, then iterates over each data row to assign a SUM($B$2:B{row}) formula in column C, recalculates the workbook, and saves it as FinancialRunningTotal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Get the first worksheet and name it
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Financial";

        // Add headers
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["C1"].PutValue("Running Total");

        // Populate sample data (rows 2 to 10)
        for (int i = 2; i <= 10; i++)
        {
            // Date column (A)
            sheet.Cells[i, 0].PutValue(DateTime.Today.AddDays(i - 2));
            // Amount column (B) – example values
            sheet.Cells[i, 1].PutValue(i * 100);
        }

        // Insert running total formula in column C for each data row
        // Formula for row r: =SUM($B$2:B{r})
        for (int row = 2; row <= sheet.Cells.MaxDataRow; row++)
        {
            string formula = $"SUM($B$2:B{row})";
            sheet.Cells[row, 2].Formula = formula;
        }

        // Recalculate all formulas so the running totals are evaluated
        workbook.CalculateFormula();

        // Save the workbook (lifecycle rule)
        workbook.Save("FinancialRunningTotal.xlsx");
    }
}
