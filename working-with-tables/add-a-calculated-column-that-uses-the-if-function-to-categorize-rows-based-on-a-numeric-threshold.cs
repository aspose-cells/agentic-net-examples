// Title: Create an Excel file with a calculated column using IF to label values above or below a threshold in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds a workbook, writes numeric data to column A, inserts an IF expression in column B that returns "Above" when the value exceeds a given threshold and "Below" otherwise, then evaluates all formulas. | Write a C# snippet using Aspose.Cells to apply the same conditional expression to each row dynamically, compute the results, and save the workbook as an .xlsx file.
// Common Searches: asp.net how to insert a conditional expression in each cell of a column with Aspose.Cells | c# Aspose.Cells calculate column based on numeric threshold | using Aspose.Cells to categorize rows as above or below a value | apply conditional formula to a range of cells in Aspose.Cells C# | evaluate formulas after adding IF column in Aspose.Cells workbook
// Tags: Aspose.Cells insert conditional expression | C# compute workbook calculations | Aspose.Cells create calculated column | Excel numeric threshold labeling | Aspose.Cells export XLSX file

using Aspose.Cells;
using System;

// // This program creates a new workbook, writes a set of numeric values to column A, adds an IF expression in column B that marks each value as "Above" or "Below" depending on a 60‑point threshold, calculates all formulas, and saves the result as CalculatedColumn.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add headers
        sheet.Cells["A1"].PutValue("Value");
        sheet.Cells["B1"].PutValue("Category");

        // Sample numeric data in column A
        double[] values = { 45, 78, 30, 90, 55 };
        for (int i = 0; i < values.Length; i++)
        {
            // Row index is zero‑based; i+1 corresponds to Excel rows 2,3,...
            sheet.Cells[i + 1, 0].PutValue(values[i]); // Column A
        }

        // Numeric threshold for categorization
        double threshold = 60;

        // Add IF formula in column B to categorize each row
        for (int i = 0; i < values.Length; i++)
        {
            int excelRow = i + 2; // Excel row number (2,3,...)
            // Formula: IF(A{row}>threshold,"Above","Below")
            string formula = $"IF(A{excelRow}>{threshold},\"Above\",\"Below\")";
            sheet.Cells[i + 1, 1].Formula = formula; // Column B
        }

        // Evaluate all formulas
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("CalculatedColumn.xlsx");
    }
}
