using Aspose.Cells;
using System;

class FindFormulaReferences
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some data in column A
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].PutValue(30);

        // Define a named range "MyRange" that refers to A1:A3
        int nameIndex = workbook.Worksheets.Names.Add("MyRange");
        workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$A$3";

        // Add formulas that reference the named range
        worksheet.Cells["B1"].Formula = "=SUM(MyRange)";
        worksheet.Cells["B2"].Formula = "=AVERAGE(MyRange)";
        // This cell does NOT reference the named range (for contrast)
        worksheet.Cells["C1"].Formula = "=A1*2";

        // Optional: calculate formulas so that values are up‑to‑date
        workbook.CalculateFormula();

        // Configure FindOptions to search only within formulas and look for the name text
        FindOptions findOptions = new FindOptions
        {
            LookInType = LookInType.OnlyFormulas,
            LookAtType = LookAtType.Contains
        };

        // Iterate through all cells whose formulas contain "MyRange"
        Cell previousCell = null;
        while (true)
        {
            Cell foundCell = worksheet.Cells.Find("MyRange", previousCell, findOptions);
            if (foundCell == null)
                break;

            Console.WriteLine($"Found reference in cell {foundCell.Name}: Formula = {foundCell.Formula}");
            previousCell = foundCell; // continue searching after the current cell
        }

        // Save the workbook (optional)
        workbook.Save("FindFormulaReferences.xlsx");
    }
}