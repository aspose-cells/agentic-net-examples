// Title: C# – Add a calculated column with IF formula to classify rows by threshold using Aspose.Cells
// Description: Creates a new workbook, writes numeric values to column A, defines a threshold, inserts an IF formula in column B that returns “High” when the value exceeds the threshold and “Low” otherwise, calculates all formulas, and saves the file as CalculatedColumn.xlsx.
// Keywords: Aspose.Cells | C# | calculated column | IF formula | threshold classification | programmatic Excel | formula calculation | Excel automation | data categorization
// Common Searches: Aspose.Cells add IF formula column | C# set Excel formula programmatically | categorize rows by value Aspose.Cells | recalculate formulas after inserting Aspose.Cells | create calculated column .NET Excel library
// Developer Intent: Generate a worksheet, populate numeric data, and programmatically add a calculated column that labels each row as High or Low based on a defined numeric threshold.
// Use Cases: Automatically flag values that exceed a limit for quick review in generated reports. | Provide a derived classification column for downstream conditional formatting or pivot tables. | Export raw data with an added category column without manual Excel editing.
// AI Prompts: Write C# code with Aspose.Cells to add a calculated column that marks values above 75 as "Pass" and others as "Fail". | Show how to use a lookup table to apply a different threshold per row in an Aspose.Cells workbook. | Explain how to force a full recalculation of all formulas after updating cell values in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCalculatedColumnDemo
{
    // Creates a new workbook, writes numeric values to column A, defines a threshold, inserts an IF formula in column B that returns “High” when the value exceeds the threshold and “Low” otherwise, calculates all formulas, and saves the file as CalculatedColumn.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add headers for the data and the calculated column
            cells["A1"].PutValue("Value");      // Original numeric data
            cells["B1"].PutValue("Category");   // Calculated column header

            // Populate sample numeric data in column A (rows 2 to 6)
            double[] sampleValues = { 30, 55, 20, 80, 45 };
            for (int i = 0; i < sampleValues.Length; i++)
            {
                // Row index is i + 2 because Excel rows are 1‑based and we start after the header
                cells[i + 1, 0].PutValue(sampleValues[i]); // Column A (index 0)
            }

            // Define the numeric threshold for categorization
            double threshold = 50;

            // Add the IF formula to the first cell of the calculated column (B2)
            // The formula will be copied to the rest of the rows programmatically
            for (int row = 1; row <= sampleValues.Length; row++)
            {
                // Build the formula string for the current row, e.g. =IF(A2>50,"High","Low")
                string formula = $"=IF(A{row + 1}>{threshold},\"High\",\"Low\")";
                cells[row, 1].Formula = formula; // Column B (index 1)
            }

            // Optionally calculate all formulas so that the workbook stores the results
            workbook.CalculateFormula();

            // Save the workbook to a file
            workbook.Save("CalculatedColumn.xlsx");
        }
    }
}
