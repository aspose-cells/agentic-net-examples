using System;
using Aspose.Cells;

class MinMaxIfsDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook (create if not present)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data (optional – ensures the demo works even on an empty file)
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue(30);
        cells["A4"].PutValue(40);
        cells["A5"].PutValue(50);

        cells["B1"].PutValue("X");
        cells["B2"].PutValue("Y");
        cells["B3"].PutValue("X");
        cells["B4"].PutValue("Y");
        cells["B5"].PutValue("X");

        // Define MINIFS and MAXIFS formulas
        string minifsFormula = "=MINIFS(A1:A5,B1:B5,\"X\")";
        string maxifsFormula = "=MAXIFS(A1:A5,B1:B5,\"X\")";

        // Calculate the formulas directly without writing them to cells
        object minResult = worksheet.CalculateFormula(minifsFormula);
        object maxResult = worksheet.CalculateFormula(maxifsFormula);

        // Output the calculated results
        Console.WriteLine("MINIFS result (minimum A where B = \"X\"): " + minResult);
        Console.WriteLine("MAXIFS result (maximum A where B = \"X\"): " + maxResult);

        // Optionally write the formulas to cells so the results are stored in the workbook
        cells["D1"].Formula = minifsFormula;
        cells["D2"].Formula = maxifsFormula;

        // Recalculate all formulas in the workbook to populate D1 and D2
        workbook.CalculateFormula();

        // Save the workbook with the new formulas and results
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}