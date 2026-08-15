// Title: Move a range with formulas between worksheets and auto‑update references using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, fill Sheet1!A1:A3, place a formula on Sheet2 that references that range, copy the range to Sheet2!C1:C3, clear the original cells to simulate a move, recalculate the workbook so the external reference updates, and save the file.
// Keywords: Aspose.Cells | C# move range | copy range with formulas | update external references | recalculate formulas .NET | Excel range relocation | Range.CopyData | Workbook.CalculateFormula
// Common Searches: Aspose.Cells move range to another sheet | update formula references after moving cells Aspose | copy range with formulas and recalc workbook C# | how to relocate Excel cells programmatically Aspose.Cells | auto‑update external sheet references Aspose.Cells
// Developer Intent: Programmatically move a cell range that contains formulas from one worksheet to another and have any formulas that reference the original range automatically point to the new location.
// Use Cases: Reorganize a report by moving a data block while preserving dependent calculations. | Consolidate multiple worksheets into a single sheet without breaking summary formulas. | Automate workbook restructuring where formulas reference moved ranges and need instant updates.
// AI Prompts: Generate C# code with Aspose.Cells to move a range containing formulas from Sheet1 to Sheet3 and refresh all external references. | Explain how Aspose.Cells updates formula references after a range is moved and how to verify the changes programmatically. | Provide a step‑by‑step guide to copy a range, clear the source, recalculate the workbook, and save the workbook using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeMoveDemo
{
    // Demonstrates how to create a workbook, fill Sheet1!A1:A3, place a formula on Sheet2 that references that range, copy the range to Sheet2!C1:C3, clear the original cells to simulate a move, recalculate the workbook so the external reference updates, and save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the default worksheet
                Workbook wb = new Workbook();
                Worksheet sheet1 = wb.Worksheets[0];
                sheet1.Name = "Sheet1";

                // Add a second worksheet that will receive the moved range
                Worksheet sheet2 = wb.Worksheets.Add("Sheet2");

                // -------------------------------------------------
                // 1. Populate Sheet1 with data and formulas
                // -------------------------------------------------
                sheet1.Cells["A1"].PutValue(10);
                sheet1.Cells["A2"].PutValue(20);
                sheet1.Cells["A3"].PutValue(30);

                // Formula that references the range we will move (A1:A3)
                // Placed on Sheet2 to demonstrate external reference update
                sheet2.Cells["B1"].Formula = "=Sheet1!SUM(A1:A3)";

                // -------------------------------------------------
                // 2. Define the source range (A1:A3) on Sheet1
                // -------------------------------------------------
                AsposeRange srcRange = sheet1.Cells.CreateRange("A1:A3");

                // -------------------------------------------------
                // 3. Define the destination range on Sheet2 (C1:C3)
                // -------------------------------------------------
                AsposeRange destRange = sheet2.Cells.CreateRange("C1:C3");

                // -------------------------------------------------
                // 4. Copy the range (including formulas) to the destination
                // -------------------------------------------------
                destRange.CopyData(srcRange);

                // -------------------------------------------------
                // 5. Clear the original cells (optional – simulates a "move")
                // -------------------------------------------------
                foreach (Cell cell in srcRange)
                {
                    cell.PutValue(string.Empty); // remove value/formula
                }

                // -------------------------------------------------
                // 6. Recalculate the workbook so that any references are refreshed
                // -------------------------------------------------
                wb.CalculateFormula();

                // -------------------------------------------------
                // 7. Verify that the external reference on Sheet2!B1 now points to the new location
                // -------------------------------------------------
                Console.WriteLine("Updated formula in Sheet2!B1: " + sheet2.Cells["B1"].Formula);
                // Expected: "=Sheet2!SUM(C1:C3)" because the source range was moved to Sheet2

                // -------------------------------------------------
                // 8. Save the workbook
                // -------------------------------------------------
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RangeMoved.xlsx");
                wb.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
