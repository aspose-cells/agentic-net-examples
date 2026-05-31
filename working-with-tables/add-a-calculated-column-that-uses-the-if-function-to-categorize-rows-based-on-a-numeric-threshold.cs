using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample numeric data in column A
        cells["A1"].PutValue("Value");
        cells["A2"].PutValue(50);
        cells["A3"].PutValue(150);
        cells["A4"].PutValue(80);
        cells["A5"].PutValue(200);

        // Add header for the calculated column in column B
        cells["B1"].PutValue("Category");

        // Define the numeric threshold for categorization
        double threshold = 100;

        // Apply IF formula to each row to categorize based on the threshold
        for (int row = 2; row <= 5; row++)
        {
            // Formula: =IF(A{row}>threshold,"High","Low")
            string formula = $"=IF(A{row}>{threshold},\"High\",\"Low\")";
            cells[$"B{row}"].Formula = formula;
        }

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("CalculatedColumn.xlsx");
    }
}