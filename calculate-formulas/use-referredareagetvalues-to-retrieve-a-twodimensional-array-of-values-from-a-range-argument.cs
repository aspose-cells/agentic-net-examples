using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate a 2x3 range with sample data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["B1"].PutValue(30);
        sheet.Cells["B2"].PutValue(40);
        sheet.Cells["C1"].PutValue(50);
        sheet.Cells["C2"].PutValue(60);

        // Set a formula that references the range A1:C2
        Cell formulaCell = sheet.Cells["E1"];
        formulaCell.Formula = "=SUM(A1:C2)";

        // Calculate formulas so that precedents are up‑to‑date
        workbook.CalculateFormula();

        // Retrieve the referred areas of the formula cell
        ReferredAreaCollection precedents = formulaCell.GetPrecedents();

        if (precedents != null && precedents.Count > 0)
        {
            // Use the first referred area (the range A1:C2)
            ReferredArea area = precedents[0];

            // Get all values in the area as a 2‑dimensional array
            object values = area.GetValues();

            // When the area contains more than one cell, GetValues returns object[,]
            if (values is object[,] multiArray)
            {
                Console.WriteLine("Values in the referred area:");
                for (int r = 0; r < multiArray.GetLength(0); r++)
                {
                    for (int c = 0; c < multiArray.GetLength(1); c++)
                    {
                        Console.Write(multiArray[r, c] + "\t");
                    }
                    Console.WriteLine();
                }
            }
            else
            {
                // Single‑cell area returns a scalar value
                Console.WriteLine("Single value: " + values);
            }
        }

        // Save the workbook (save rule)
        workbook.Save("ReferredAreaValuesDemo.xlsx");
    }
}