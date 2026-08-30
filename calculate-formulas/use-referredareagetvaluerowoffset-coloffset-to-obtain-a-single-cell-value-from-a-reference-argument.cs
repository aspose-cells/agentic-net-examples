// Title: How to read a single cell value from a formula's precedent range using ReferredArea.GetValue in Aspose.Cells for .NET
// AI Prompts: Write C# code that obtains the first ReferredArea of a formula cell and calls ReferredArea.GetValue(rowOffset, colOffset) to fetch the value of a specific referenced cell. | Demonstrate iterating over row and column offsets within a ReferredArea to extract values from any cell in the range referenced by a SUM formula.
// Common Searches: Aspose.Cells C# get value from referenced range of a formula | How to use ReferredArea.GetValue to read a cell at a given offset in .NET | Retrieve precedent cell values from a SUM formula using Aspose.Cells API | Example of GetPrecedents and GetValue methods in Aspose.Cells for C#
// Tags: ReferredArea.GetValue offset extraction | GetPrecedents retrieve formula precedents Aspose.Cells | Aspose.Cells read specific cell from range | C# Aspose.Cells formula precedent value | Aspose.Cells workbook calculation and cell access

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, places numbers in cells A1:B2, assigns a SUM(A1:B2) formula to C1, forces formula calculation, obtains the formula's precedent range via GetPrecedents, and uses ReferredArea.GetValue with zero‑based row and column offsets to read the values at A1 and B2, printing them before saving the workbook.
    public class ReferredAreaGetValueDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put some values in cells that will be referenced
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["B1"].PutValue(30);
                sheet.Cells["B2"].PutValue(40);

                // Create a formula that references the range A1:B2
                Cell formulaCell = sheet.Cells["C1"];
                formulaCell.Formula = "=SUM(A1:B2)";

                // Ensure all formulas are calculated so that precedents are resolved
                workbook.CalculateFormula();

                // Get the collection of referred areas (precedents) for the formula cell
                ReferredAreaCollection precedents = formulaCell.GetPrecedents();

                if (precedents != null && precedents.Count > 0)
                {
                    // Take the first referred area (in this case the range A1:B2)
                    ReferredArea area = precedents[0];

                    // Obtain a single cell value from the area using row and column offsets
                    // Offsets are zero‑based from the top‑left cell of the area.
                    // Example: (0,0) -> A1, (0,1) -> B1, (1,0) -> A2, (1,1) -> B2
                    object valueAtA1 = area.GetValue(0, 0);
                    object valueAtB2 = area.GetValue(1, 1);

                    Console.WriteLine($"Value at area offset (0,0) [{CellsHelper.CellIndexToName(area.StartRow, area.StartColumn)}]: {valueAtA1}");
                    Console.WriteLine($"Value at area offset (1,1) [{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}]: {valueAtB2}");
                }
                else
                {
                    Console.WriteLine("No precedents found for the formula cell.");
                }

                // Save the workbook (optional, demonstrates lifecycle compliance)
                workbook.Save("ReferredAreaGetValueDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime exception: {ex.Message}");
                throw;
            }
        }
    }
}
