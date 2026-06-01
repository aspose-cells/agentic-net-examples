using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data that the formula will reference
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].PutValue(30);
        worksheet.Cells["B1"].PutValue(2);
        worksheet.Cells["B2"].PutValue(4);
        worksheet.Cells["B3"].PutValue(6);

        // Define a complex formula that uses the data above
        // Example: (SUM of A1:A3) multiplied by (AVERAGE of B1:B3)
        string formula = "=SUM(A1:A3)*AVERAGE(B1:B3)";

        // Evaluate the formula directly without writing it to a cell
        // This uses Worksheet.CalculateFormula(string) and does not recalculate the whole sheet
        object result = worksheet.CalculateFormula(formula);

        // Output the calculated result
        Console.WriteLine($"Result of formula \"{formula}\" = {result}");

        // (Optional) Store the result in a cell for verification
        worksheet.Cells["C1"].PutValue(result);

        // Save the workbook to a file
        workbook.Save("ComplexFormulaResult.xlsx");
    }
}