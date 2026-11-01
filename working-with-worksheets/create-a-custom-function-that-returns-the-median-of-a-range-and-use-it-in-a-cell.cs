// Title: How to calculate the median of a range and store the result in another cell using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example applying MEDIAN function to a cell range | Set formula and evaluate median value in an Excel workbook with Aspose.Cells | Retrieve calculated median from a cell after calling CalculateFormula in Aspose.Cells | Save Excel file after formula evaluation using Aspose.Cells C# API | Populate worksheet with numeric data and compute median using Aspose.Cells
// Tags: Aspose.Cells MEDIAN formula C# | evaluate workbook formulas Aspose.Cells | write calculated result to cell Aspose.Cells | save Excel workbook after formula calculation Aspose.Cells | populate worksheet range with numeric data Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomMedian
{
    // The sample creates a new Workbook, fills cells A1:A5 with numbers, assigns the built‑in MEDIAN formula to cell B1, forces calculation with CalculateFormula(), reads the resulting median value, and saves the workbook as CustomMedianResult.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Fill some sample data in A1:A5.
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);
                sheet.Cells["A4"].PutValue(40);
                sheet.Cells["A5"].PutValue(50);

                // Use the built‑in MEDIAN function in cell B1.
                sheet.Cells["B1"].Formula = "MEDIAN(A1:A5)";

                // Calculate the workbook to evaluate the formula.
                workbook.CalculateFormula();

                // Retrieve and display the result.
                double medianResult = sheet.Cells["B1"].DoubleValue;
                Console.WriteLine($"Median calculated by built‑in function: {medianResult}");

                // Define output file path.
                string outputPath = "CustomMedianResult.xlsx";

                // Ensure the directory exists before saving.
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
