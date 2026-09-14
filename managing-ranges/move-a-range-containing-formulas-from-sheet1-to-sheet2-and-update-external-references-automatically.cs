// Title: Copy a formula‑filled range from Sheet1 to Sheet2 and automatically refresh external references with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to copy a cell range that contains formulas from one worksheet to another, then clear the source range and recalculate the workbook. | Programmatically move a block of formula cells between sheets in a .xlsx file and ensure all external references are updated by invoking CalculateFormula.
// Common Searches: Aspose.Cells C# copy range with formulas and recalculate after moving | How to transfer a formula block from Sheet1 to Sheet2 using Aspose.Cells | Refresh external references after moving formula range in .NET workbook
// Tags: copy formula range between sheets Aspose.Cells | clear source cells after copy C# | calculate workbook formulas Aspose.Cells | move cell block with external references .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads 'input.xlsx', copies the A1:C10 range containing formulas from Sheet1 to Sheet2, clears the original cells, recalculates all formulas to update external references, and saves the result as 'output.xlsx' using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Get source and destination worksheets
                var sheet1 = workbook.Worksheets["Sheet1"];
                var sheet2 = workbook.Worksheets["Sheet2"];

                if (sheet1 == null || sheet2 == null)
                {
                    Console.WriteLine("Required worksheets (Sheet1/Sheet2) are missing.");
                    return;
                }

                // Define the source range that contains formulas (e.g., A1:C10)
                int srcFirstRow = 0;      // zero‑based index for row 1 (A1)
                int srcFirstColumn = 0;   // zero‑based index for column A
                int totalRows = 10;
                int totalColumns = 3;

                // Define the top‑left cell in the destination sheet where the range will be moved
                int destFirstRow = 0;     // start at A1 in Sheet2
                int destFirstColumn = 0;

                // Create the source range
                var srcRange = sheet1.Cells.CreateRange(srcFirstRow, srcFirstColumn, totalRows, totalColumns);

                // Create the destination range
                var destRange = sheet2.Cells.CreateRange(destFirstRow, destFirstColumn, totalRows, totalColumns);

                // Copy the range to the destination sheet preserving formulas and values
                srcRange.Copy(destRange);

                // Clear the original range from Sheet1 (optional, effectively "moving")
                srcRange.Clear();

                // Recalculate all formulas so that any external references are updated
                workbook.CalculateFormula();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
