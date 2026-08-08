// Title: Map Quarterly Results to a Pre‑Formatted Financial Statement with Aspose.Cells Range.CopyValue (C#)
// Description: Creates a new workbook, defines a financial‑statement template with quarter headers and row labels, loads a List<double[]> of revenue, cost and profit into a hidden source range (A11), copies the values to the visible range B2:E4 using Range.CopyValue, and saves the file as FinancialStatement.xlsx.
// Keywords: Aspose.Cells Range.CopyValue C# | populate financial statement Aspose.Cells | copy hidden source range to visible range | map List<double[]> to Excel cells | quarterly report automation Aspose | C# Excel data mapping Aspose.Cells | financial report generation .NET
// Common Searches: Aspose.Cells copy values from hidden range C# | How to map a collection to a pre‑formatted Excel layout using Aspose.Cells | Range.CopyValue example for financial statements | C# load List<double[]> into Excel with Aspose.Cells | Create and hide source range Aspose.Cells smart markers
// Developer Intent: Copy a collection of quarterly financial results into a pre‑designed Excel statement by using a hidden source range and the Range.CopyValue method.
// Use Cases: Automatically fill quarterly revenue, cost, and profit rows in a standard financial‑statement template. | Reuse a single hidden data block to populate multiple report sections without manual cell references. | Integrate dynamic Excel generation into reporting pipelines that consume in‑memory collections.
// AI Prompts: Write C# code that uses Aspose.Cells Range.CopyValue to transfer a List<double[]> into a predefined financial‑statement layout. | Explain how to create a hidden source range, populate it with data, and copy it to a visible destination range using Aspose.Cells in .NET. | Suggest performance‑optimised patterns and error‑handling best practices for copying large ranges with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace FinancialStatementMapping
{
    // Creates a new workbook, defines a financial‑statement template with quarter headers and row labels, loads a List<double[]> of revenue, cost and profit into a hidden source range (A11), copies the values to the visible range B2:E4 using Range.CopyValue, and saves the file as FinancialStatement.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Define the layout of the financial statement
                //    Row 1: Headers for quarters
                cells["B1"].PutValue("Q1");
                cells["C1"].PutValue("Q2");
                cells["D1"].PutValue("Q3");
                cells["E1"].PutValue("Q4");

                //    Column A: Row labels
                cells["A2"].PutValue("Revenue");
                cells["A3"].PutValue("Cost");
                cells["A4"].PutValue("Profit");

                // 3. Prepare a collection of quarterly results.
                //    Each inner array represents a row (Revenue, Cost, Profit) for Q1‑Q4.
                List<double[]> quarterlyResults = new List<double[]>
                {
                    new double[] { 150000, 180000, 210000, 240000 }, // Revenue
                    new double[] {  90000, 110000, 130000, 150000 }, // Cost
                    new double[] {  60000,  70000,  80000,  90000 }  // Profit
                };

                // 4. Create a hidden source range where we will load the collection.
                //    Start at row 11 (index 10) column A (index 0) – this area is not visible in the final report.
                int srcStartRow = 10;   // zero‑based index
                int srcStartCol = 0;
                int rowCount = quarterlyResults.Count;      // 3 rows
                int colCount = quarterlyResults[0].Length; // 4 columns

                AsposeRange srcRange = cells.CreateRange(srcStartRow, srcStartCol, rowCount, colCount);

                // Fill the source range with data from the collection.
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        srcRange[r, c].PutValue(quarterlyResults[r][c]);
                    }
                }

                // 5. Define the destination range that matches the pre‑formatted layout.
                //    It starts at cell B2 (row index 1, column index 1) and has the same dimensions.
                int destStartRow = 1; // B2 row
                int destStartCol = 1; // B2 column
                AsposeRange destRange = cells.CreateRange(destStartRow, destStartCol, rowCount, colCount);

                // 6. Map the source range into the destination range using the Range.CopyValue method.
                destRange.CopyValue(srcRange);

                // 7. Save the workbook (lifecycle rule: save)
                string outputPath = "FinancialStatement.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
