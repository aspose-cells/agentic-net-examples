// Title: Aspose.Cells C# – Retrieve a Single Precedent Cell Value Using ReferredArea.GetValue
// Description: Demonstrates how to calculate formulas, obtain a cell's precedents with GetPrecedents, and read the referenced value via ReferredArea.GetValue(0,0) in a .NET workbook.
// Keywords: Aspose.Cells ReferredArea.GetValue | C# get precedent cell value | Aspose.Cells formula precedents | Read referenced cell Aspose.Cells | GetPrecedents example .NET | Aspose.Cells offset cell retrieval
// Common Searches: how to read a precedent cell in Aspose.Cells | Aspose.Cells GetValue row offset column offset | C# example for GetPrecedents and GetValue | retrieve formula reference value Aspose.Cells | Aspose.Cells single cell value from referenced area
// Developer Intent: Extract the value of a cell referenced by a formula by accessing its ReferredArea and calling GetValue with the appropriate offsets.
// Use Cases: Validate a precedent cell's content before performing custom business logic. | Log or audit all cells that a formula depends on for debugging complex spreadsheets. | Iterate over a multi‑cell precedent range and collect each cell's value using different offsets.
// AI Prompts: Write C# code that takes a formula cell, gets its first ReferredArea, and returns the value at offset (0,0) with error handling for missing precedents. | Create a method that enumerates all ReferredArea objects of a formula cell and builds a dictionary of (rowOffset, colOffset) → cell value using GetValue. | Show how to safely call ReferredArea.GetValue only when the requested offset exists, returning null or a default value otherwise.

using System;
using Aspose.Cells;

namespace ReferredAreaGetValueDemo
{
    // Demonstrates how to calculate formulas, obtain a cell's precedents with GetPrecedents, and read the referenced value via ReferredArea.GetValue(0,0) in a .NET workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells that will be referenced by a formula
            sheet.Cells["A1"].PutValue(10);   // First value
            sheet.Cells["B1"].PutValue(20);   // Second value
            sheet.Cells["A2"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);

            // Set a formula that references cell A1
            Cell formulaCell = sheet.Cells["C1"];
            formulaCell.Formula = "=A1";

            // Calculate formulas so that the referenced value is up‑to‑date
            workbook.CalculateFormula();

            // Get the collection of areas that the formula cell depends on
            ReferredAreaCollection precedents = formulaCell.GetPrecedents();

            if (precedents != null && precedents.Count > 0)
            {
                // Take the first referred area (in this case it is a single cell A1)
                ReferredArea area = precedents[0];

                // Obtain the value at offset (0,0) – the top‑left cell of the area
                object valueAt00 = area.GetValue(0, 0);
                Console.WriteLine($"Value at offset (0,0): {valueAt00}");

                // If the area were larger, you could retrieve other cells by changing offsets
                // Example: get value at row offset 0, column offset 1 (B1) – only works if the area includes it
                // object valueAt01 = area.GetValue(0, 1);
                // Console.WriteLine($"Value at offset (0,1): {valueAt01}");
            }
            else
            {
                Console.WriteLine("No precedents found for the formula cell.");
            }

            // Save the workbook (optional, demonstrates lifecycle compliance)
            workbook.Save("ReferredAreaGetValueDemo.xlsx");
        }
    }
}
