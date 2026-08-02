// Title: Create a column‑wide named range and compute its average with Aspose.Cells for .NET (C#)
// Description: This C# example shows how to generate a new workbook, fill column B with numbers, use the EntireColumn property to build a named range that covers the whole column, apply =AVERAGE(MyColumn) in cell C1, evaluate the formula, output the result, and save the file as NamedColumnAverage.xlsx.
// Keywords: Aspose.Cells named range column | C# entire column reference | average formula Aspose.Cells | column‑wide named range .NET | Workbook.CalculateFormula | CreateRange EntireColumn | Aspose.Cells GitHub example | Excel column average C#
// Common Searches: Aspose.Cells create named range for whole column | C# calculate average of a column using named range | How to use EntireColumn property in Aspose.Cells | Define column‑wide named range in .NET | Aspose.Cells average formula example
// Developer Intent: Define a named range that spans an entire worksheet column and use it in a formula to calculate the column’s average.
// Use Cases: Generate a reusable column‑wide named range for summary calculations across multiple sheets. | Insert =AVERAGE(MyColumn) into any cell to display the average of the referenced column. | Automate reporting by saving the workbook with the computed average and the named range for future reference.
// AI Prompts: Write C# code with Aspose.Cells that creates a named range for column D and returns its sum using =SUM(). | Explain how the EntireColumn property can be leveraged to build dynamic column‑wide named ranges for formulas. | Provide error‑handling strategies when the column‑wide named range contains mixed data types while calculating an average.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedColumnAverage
{
    // This C# example shows how to generate a new workbook, fill column B with numbers, use the EntireColumn property to build a named range that covers the whole column, apply =AVERAGE(MyColumn) in cell C1, evaluate the formula, output the result, and save the file as NamedColumnAverage.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample numeric data in column B (index 1)
                for (int i = 0; i < 10; i++)
                {
                    cells[i, 1].PutValue(i + 1); // B1..B10 = 1..10
                }

                // Create a temporary range that starts at B1
                // Then obtain the entire column that contains this range
                Aspose.Cells.Range tempRange = cells.CreateRange(0, 1, 1, 1);
                Aspose.Cells.Range entireColumn = tempRange.EntireColumn;

                // Define a named range that refers to the whole column B
                int nameIndex = workbook.Worksheets.Names.Add("MyColumn");
                Name namedRange = workbook.Worksheets.Names[nameIndex];

                // Build the reference string like =Sheet1!$B:$B
                string columnLetter = CellsHelper.ColumnIndexToName(1); // column B
                namedRange.RefersTo = $"={sheet.Name}!${columnLetter}:${columnLetter}";

                // Use the named range in a formula to calculate the average
                // (use the literal name to avoid potential property issues)
                cells["C1"].Formula = $"=AVERAGE(MyColumn)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Output the result to the console
                Console.WriteLine($"Average of column {columnLetter}: {cells["C1"].Value}");

                // Save the workbook
                string outputPath = "NamedColumnAverage.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
