// Title: Assign the same formula to each cell in column B with a relative reference to column A using Aspose.Cells for .NET
// Common Searches: how to set the same formula for every cell in column B using Aspose.Cells in C# | assign a formula that references column A for each row with Aspose.Cells .NET loop | apply =A1+5 formula to column B programmatically with Aspose.Cells | recalculate all formulas after assigning them with Aspose.Cells C# | check calculated values of column B after using workbook.CalculateFormula in Aspose.Cells
// Tags: apply formula to column Aspose.Cells C# | relative cell reference in Aspose.Cells formula | loop assign formulas Aspose.Cells .NET | recalculate workbook formulas Aspose.Cells | verify formula results Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, fills column A with numbers 1‑10, uses a C# loop to set the formula "=A{row}+5" in each cell of column B, recalculates the workbook, prints and validates the results, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill column A with sample data (1 to 10)
            for (int row = 0; row < 10; row++)
            {
                cells[row, 0].PutValue(row + 1); // A column
            }

            // Assign the same formula to each cell in column B using a loop
            // The formula adds 5 to the value in column A of the same row.
            for (int row = 0; row < 10; row++)
            {
                Cell cellB = cells[row, 1]; // B column
                // Use A1 notation with the specific row number
                cellB.Formula = $"=A{row + 1}+5";
            }

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Verify that the relative references produced the expected results
            for (int row = 0; row < 10; row++)
            {
                double valueA = cells[row, 0].DoubleValue;
                double valueB = cells[row, 1].DoubleValue;
                Console.WriteLine($"Row {row + 1}: A = {valueA}, B = {valueB} (expected {valueA + 5})");
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputPath = "Result.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
