// Title: How to use ReferredArea.GetValues to read a 2‑D array of cell values from a formula's precedent range in Aspose.Cells for .NET
// AI Prompts: Generate C# code that obtains the first ReferredArea of a cell containing a SUM formula and calls GetValues() to return the range values as an object[,] array. | Write a C# snippet that iterates over the object[][] jagged array returned by ReferredArea.GetValues() and prints each cell value. | Provide an example that calculates formulas, extracts the precedent values, and then saves the workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# get values from a referenced range using ReferredArea.GetValues | How to retrieve a two dimensional array from a formula precedent in .NET | Example of object[,] vs object[][] returned by ReferredArea.GetValues in Aspose.Cells | Read cell values of SUM formula range programmatically with Aspose.Cells | Calculate formulas then access precedent cell values Aspose.Cells C#
// Tags: Aspose.Cells ReferredArea.GetValues example | retrieve 2D array from precedent range C# | calculate formulas then read cell values Aspose.Cells | object[,] handling in Aspose.Cells | object[][] handling in Aspose.Cells

using System;
using Aspose.Cells;

namespace ReferredAreaGetValuesDemo
{
    // The sample creates a workbook, fills cells A1:B2 with numbers, sets C1 to =SUM(A1:B2), calculates formulas, obtains the first ReferredArea from C1's precedents, calls GetValues() to retrieve the range values as either an object[,] or object[][] array, iterates and prints the values, and finally saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a 2x2 range with sample data
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["B2"].PutValue(4);

            // Create a formula that references the range A1:B2
            Cell formulaCell = sheet.Cells["C1"];
            formulaCell.Formula = "=SUM(A1:B2)";

            // Calculate formulas so that any dependent values are up‑to‑date
            workbook.CalculateFormula();

            // Get the collection of areas that the formula cell refers to
            ReferredAreaCollection precedents = formulaCell.GetPrecedents();

            if (precedents != null && precedents.Count > 0)
            {
                // Take the first referred area (A1:B2 in this case)
                ReferredArea area = precedents[0];

                // Retrieve all values from the area
                object valuesObj = area.GetValues();

                // The returned object can be a 2‑D array (object[,]) or a jagged array (object[][])
                if (valuesObj is object[,] multiArray)
                {
                    Console.WriteLine("Values retrieved via object[,]:");
                    for (int r = 0; r < multiArray.GetLength(0); r++)
                    {
                        for (int c = 0; c < multiArray.GetLength(1); c++)
                        {
                            Console.Write(multiArray[r, c] + "\t");
                        }
                        Console.WriteLine();
                    }
                }
                else if (valuesObj is object[][] jaggedArray)
                {
                    Console.WriteLine("Values retrieved via object[][]:");
                    foreach (object[] row in jaggedArray)
                    {
                        foreach (object val in row)
                        {
                            Console.Write(val + "\t");
                        }
                        Console.WriteLine();
                    }
                }
                else
                {
                    // Single cell case
                    Console.WriteLine("Single value: " + valuesObj);
                }
            }

            // Save the workbook (optional, just to demonstrate the save rule)
            workbook.Save("ReferredAreaGetValuesDemo.xlsx");
        }
    }
}
