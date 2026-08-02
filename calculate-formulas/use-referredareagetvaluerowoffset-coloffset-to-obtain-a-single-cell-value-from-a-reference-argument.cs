// Title: Aspose.Cells C# – Retrieve a Single Cell Value from a Formula’s Referenced Range with ReferredArea.GetValue
// Description: This example creates a workbook, fills cells A1‑B2, assigns a SUM formula to C1, calculates all formulas, obtains the precedent range via GetPrecedents, and uses ReferredArea.GetValue(0,1) to read the value of cell B1 inside that range. The retrieved value is printed and the workbook is saved.
// Keywords: Aspose.Cells ReferredArea.GetValue | C# get cell value from formula precedent | Aspose.Cells GetPrecedents example | read cell by offset Aspose.Cells | calculate formulas Aspose.Cells | extract single cell value .NET | Aspose.Cells workbook example
// Common Searches: Aspose.Cells C# get value from referenced range | How to use ReferredArea.GetValue in .NET | Read B1 from SUM(A1:B2) with Aspose.Cells | GetPrecedents and GetValue example C# | Retrieve single cell value from formula precedent
// Developer Intent: Fetch the value of a specific cell inside a formula’s referenced range using ReferredArea.GetValue after the workbook has been calculated.
// Use Cases: Debugging: display individual precedent values to verify formula inputs. | Reporting: extract particular cells from a summed range for custom summaries. | Conditional logic: feed a single precedent value into further .NET processing after calculation.
// AI Prompts: Generate C# code that uses Aspose.Cells to obtain the value of cell B2 from a SUM(A1:B2) formula with ReferredArea.GetValue. | Write a snippet that loops through all ReferredArea objects of a formula cell and prints each cell’s value using GetValue. | Explain how to cast and handle different data types returned by ReferredArea.GetValue in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, fills cells A1‑B2, assigns a SUM formula to C1, calculates all formulas, obtains the precedent range via GetPrecedents, and uses ReferredArea.GetValue(0,1) to read the value of cell B1 inside that range. The retrieved value is printed and the workbook is saved.
    public class ReferredAreaGetValueDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some cells with values
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["B1"].PutValue(30);
                sheet.Cells["B2"].PutValue(40);

                // Set a formula that references the range A1:B2
                Cell formulaCell = sheet.Cells["C1"];
                formulaCell.Formula = "=SUM(A1:B2)";

                // Calculate all formulas so that precedents are up‑to‑date
                workbook.CalculateFormula();

                // Retrieve the collection of referred areas (precedents) for the formula cell
                ReferredAreaCollection precedents = formulaCell.GetPrecedents();

                if (precedents != null && precedents.Count > 0)
                {
                    // Use the first referred area (which corresponds to A1:B2)
                    ReferredArea area = precedents[0];

                    // Get the value of the cell at row offset 0, column offset 1 within the area (i.e., B1)
                    object value = area.GetValue(0, 1);

                    Console.WriteLine($"Value at offset (0,1) within the referred area: {value}");
                }

                // Save the workbook (optional)
                workbook.Save("ReferredAreaGetValueDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ReferredAreaGetValueDemo.Run();
        }
    }
}
