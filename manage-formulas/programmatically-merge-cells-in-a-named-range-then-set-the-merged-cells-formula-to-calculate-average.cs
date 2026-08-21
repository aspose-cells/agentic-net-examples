// Title: Merge a Named Range and Apply an AVERAGE Formula with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, fills cells A1‑B2 with numbers, defines a named range, merges the range, assigns the formula =AVERAGE(A1:B2) to the upper‑left cell, recalculates formulas, and saves the file as an XLSX document.
// Keywords: Aspose.Cells merge cells C# | Aspose.Cells set formula | Aspose.Cells average calculation | Aspose.Cells named range | Aspose.Cells recalculate formulas | C# Excel merge range | Excel AVERAGE formula Aspose | Aspose.Cells workbook save
// Common Searches: how to merge a range and set a formula with Aspose.Cells .NET | Aspose.Cells C# merge cells and calculate average | programmatically merge cells and apply AVERAGE in Excel using Aspose | Aspose.Cells example for merging a named range | C# code to merge cells and recalculate formulas with Aspose.Cells
// Developer Intent: Merge a specific range and attach an AVERAGE formula to the merged cell using Aspose.Cells for .NET.
// Use Cases: Create a summary block that visually spans multiple columns while showing the average of the underlying data. | Design a report header that merges cells across a table and automatically displays the average of the data set. | Build a reusable template where grouped cells are merged and the average is computed on the fly.
// AI Prompts: Generate C# code with Aspose.Cells that merges a given address range and sets its formula to =AVERAGE of that range, then recalculates and saves the workbook. | Provide a method for Aspose.Cells that accepts a worksheet, a range address, and a formula string, merges the range, applies the formula to the top‑left cell, and updates the workbook. | Show how to merge cells A1:C3 using Aspose.Cells for .NET, assign the formula =SUM(A1:C3) to the merged cell, ensure proper recalculation, and export the result to XLSX.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeAndAverage
{
    // Creates a new workbook, fills cells A1‑B2 with numbers, defines a named range, merges the range, assigns the formula =AVERAGE(A1:B2) to the upper‑left cell, recalculates formulas, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["B1"].PutValue(30);
                cells["B2"].PutValue(40);

                // Create the range object for the address A1:B2
                Aspose.Cells.Range range = cells.CreateRange("A1:B2");

                // Merge the cells in the range
                range.Merge();

                // Set formula in the merged cell (upper‑left cell A1) to calculate the average
                cells["A1"].Formula = "=AVERAGE(A1:B2)";

                // Recalculate formulas
                workbook.CalculateFormula();

                // Define output file path
                string outputPath = "MergedAverageDemo.xlsx";

                // Ensure the output directory exists (if a directory is specified)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
