using System;
using Aspose.Cells;

class MinifsValidation
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set compatibility mode for Excel 2016 (disable older compatibility checks)
        wb.Settings.CheckCompatibility = false;

        // Access the first worksheet and its cells
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Populate sample data
        // Column A contains numeric values
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue(30);
        cells["A4"].PutValue(40);

        // Column B contains criteria strings
        cells["B1"].PutValue("X");
        cells["B2"].PutValue("Y");
        cells["B3"].PutValue("X");
        cells["B4"].PutValue("Y");

        // Apply MINIFS formula: minimum of A where B = "X"
        cells["C1"].Formula = "=MINIFS(A1:A4,B1:B4,\"X\")";

        // Calculate formulas
        wb.CalculateFormula();

        // Validate the result (expected 10)
        double result = cells["C1"].DoubleValue;
        Console.WriteLine("MINIFS result (original workbook): " + result);

        // Save the workbook
        string filePath = "MinifsValidation.xlsx";
        wb.Save(filePath);

        // Load the workbook back
        Workbook wbLoaded = new Workbook(filePath);

        // Recalculate after loading
        wbLoaded.CalculateFormula();

        // Validate the result again
        double loadedResult = wbLoaded.Worksheets[0].Cells["C1"].DoubleValue;
        Console.WriteLine("MINIFS result (loaded workbook): " + loadedResult);
    }
}