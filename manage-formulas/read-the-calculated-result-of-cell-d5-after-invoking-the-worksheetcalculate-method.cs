using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells that will be used in the formula
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(5);

            // Set a formula in D5 (e.g., sum of A1 and B1)
            sheet.Cells["D5"].Formula = "=A1+B1";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Read the calculated result of cell D5
            object calculatedValue = sheet.Cells["D5"].Value;

            // Output the result
            Console.WriteLine("Calculated value of D5: " + calculatedValue);

            // Save the workbook (lifecycle rule: save)
            string outputPath = "CalculatedResult.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}