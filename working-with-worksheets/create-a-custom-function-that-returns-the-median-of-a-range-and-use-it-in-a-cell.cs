// Title: Aspose.Cells C# Example – Compute Median of a Range with the MEDIAN Formula and Save the Workbook
// Description: This example creates a new Workbook, fills cells A1:A5 with numeric values, assigns the MEDIAN(A1:A5) formula to cell B1, forces formula evaluation, prints the calculated median, ensures the target folder exists, and saves the file as MedianResult.xlsx.
// Keywords: Aspose.Cells median C# | MEDIAN formula Aspose.Cells | calculate median .NET workbook | Excel formula evaluation Aspose | save workbook after calculation | populate cells programmatically | C# Excel automation
// Common Searches: Aspose.Cells how to use MEDIAN function in C# | calculate median of a range with Aspose.Cells | C# example for formula calculation and saving workbook | populate Excel cells and compute median using Aspose | ensure output directory exists before saving Aspose workbook
// Developer Intent: Use Aspose.Cells to compute the median of a numeric range and store the result in a worksheet cell.
// Use Cases: Generate a report that requires the median of a data column. | Programmatically evaluate Excel formulas and retrieve the result in .NET. | Create and save an Excel file after performing statistical calculations.
// AI Prompts: Show me C# code that defines a user‑defined median function in Aspose.Cells and applies it to a cell. | Provide an Aspose.Cells example that fills a range, sets the MEDIAN formula, calculates all formulas, and saves the workbook. | Explain how to verify or create the output directory before writing an Aspose.Cells workbook to disk.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomMedian
{
    // This example creates a new Workbook, fills cells A1:A5 with numeric values, assigns the MEDIAN(A1:A5) formula to cell B1, forces formula evaluation, prints the calculated median, ensures the target folder exists, and saves the file as MedianResult.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                var cells = ws.Cells;

                // Populate sample numeric data in A1:A5
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["A3"].PutValue(30);
                cells["A4"].PutValue(40);
                cells["A5"].PutValue(50);

                // Use the built‑in MEDIAN function in cell B1
                cells["B1"].Formula = "=MEDIAN(A1:A5)";

                // Calculate all formulas in the workbook
                wb.CalculateFormula();

                // Output the result of the median function
                Console.WriteLine("Median of A1:A5 = " + cells["B1"].Value);

                // Save the workbook
                string outputPath = "MedianResult.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
