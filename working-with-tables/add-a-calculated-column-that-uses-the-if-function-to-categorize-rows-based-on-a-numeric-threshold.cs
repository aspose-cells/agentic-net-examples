// Title: Create a calculated column with an IF formula to label rows as High or Low based on a numeric threshold using Aspose.Cells for .NET
// Common Searches: how to add an IF formula to each row in an Aspose.Cells worksheet using C# | Aspose.Cells C# calculate column values based on a numeric threshold | programmatically set conditional IF formulas in Excel with Aspose.Cells .NET | create a calculated column that categorizes data as High or Low in Aspose.Cells | use Aspose.Cells to apply a threshold-based IF function across rows | generate a new workbook with conditional categories using Aspose.Cells C#
// Tags: add calculated column IF Aspose.Cells | categorize rows by numeric threshold C# | set cell formula programmatically Aspose.Cells | apply conditional logic to Excel column .NET | calculate workbook formulas Aspose.Cells | threshold based IF function Excel C#

using System;
using Aspose.Cells;

// The example creates a new workbook, populates column A with sample numeric values, defines a numeric threshold, inserts an IF formula in column B for each row to label the value as "High" or "Low", evaluates all formulas, and saves the result as CalculatedColumn.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add headers
        sheet.Cells["A1"].PutValue("Value");
        sheet.Cells["B1"].PutValue("Category");

        // Populate sample numeric data in column A (rows 2 to 10)
        for (int row = 2; row <= 10; row++)
        {
            // Example values: 30, 45, 60, etc.
            double sampleValue = 20 + row * 5; // arbitrary numeric data
            sheet.Cells[row, 0].PutValue(sampleValue);
        }

        // Define the numeric threshold
        double threshold = 50;

        // Add calculated column using IF function to categorize rows
        // Formula: =IF(A2>threshold,"High","Low")
        for (int row = 2; row <= 10; row++)
        {
            // Build the formula string with the appropriate cell reference
            string formula = $"IF(A{row}>{threshold},\"High\",\"Low\")";

            // Set the formula in column B for the current row
            sheet.Cells[row, 1].Formula = formula;
        }

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("CalculatedColumn.xlsx");
    }
}
