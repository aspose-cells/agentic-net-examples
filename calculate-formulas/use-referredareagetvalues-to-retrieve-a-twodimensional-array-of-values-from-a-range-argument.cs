// Title: C# Aspose.Cells: Use ReferredArea.GetValues() to fetch a 2‑D array of cell values
// Description: Demonstrates creating a workbook, filling A1:B2, adding a SUM formula, calculating it, obtaining the precedent ReferredArea, and calling GetValues() to receive either an object[,] or object[][] array, then printing and saving the file.
// Keywords: Aspose.Cells ReferredArea GetValues | C# 2D array from range | object[,] vs object[][] Aspose | formula precedents extraction .NET | read cell values programmatically
// Common Searches: Aspose.Cells GetValues 2D array example | How to retrieve values from a referenced range in C# | object[,] object[][] return type Aspose.Cells | GetPrecedents and GetValues usage
// Developer Intent: Extract every value from the range a formula references by using ReferredArea.GetValues().
// Use Cases: Validate or transform data after a formula has been evaluated. | Export a calculated range to CSV, JSON, or another system. | Process both rectangular and jagged return structures when handling dynamic or multi‑area references.
// AI Prompts: Write a C# routine that reads a named range with ReferredArea.GetValues() and writes the result to a CSV file. | Explain how to differentiate between object[,] and object[][] returned by GetValues and handle each format. | Adapt the sample to support a formula with multiple non‑contiguous precedent areas, such as SUM(A1:A2, C1:C2).

using System;
using Aspose.Cells;

namespace ReferredAreaGetValuesDemo
{
    // Demonstrates creating a workbook, filling A1:B2, adding a SUM formula, calculating it, obtaining the precedent ReferredArea, and calling GetValues() to receive either an object[,] or object[][] array, then printing and saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a 2x2 range with sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);

            // Place a formula that references the range A1:B2
            Cell formulaCell = sheet.Cells["C1"];
            formulaCell.Formula = "=SUM(A1:B2)";

            // Ensure formulas are calculated so that the referenced area is valid
            workbook.CalculateFormula();

            // Retrieve the collection of areas that the formula cell refers to
            ReferredAreaCollection precedents = formulaCell.GetPrecedents();

            if (precedents != null && precedents.Count > 0)
            {
                // Get the first referred area (A1:B2 in this case)
                ReferredArea area = precedents[0];

                // Use GetValues() to obtain all cell values in the area
                object valuesObj = area.GetValues();

                // The returned object can be a 2‑D array (object[,]) or a jagged array (object[][])
                if (valuesObj is object[,] multiArray)
                {
                    Console.WriteLine("Values retrieved via object[,]:");
                    for (int r = 0; r < multiArray.GetLength(0); r++)
                    {
                        for (int c = 0; c < multiArray.GetLength(1); c++)
                        {
                            Console.Write($"{multiArray[r, c]}\t");
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
                            Console.Write($"{val}\t");
                        }
                        Console.WriteLine();
                    }
                }
                else
                {
                    // Single cell case
                    Console.WriteLine($"Single value: {valuesObj}");
                }
            }
            else
            {
                Console.WriteLine("No referred areas found.");
            }

            // Save the workbook (optional, just to demonstrate lifecycle compliance)
            workbook.Save("ReferredAreaGetValuesDemo.xlsx");
        }
    }
}
