// Title: C# batch script to open multiple Excel workbooks, apply a shared formula, and save using Aspose.Cells
// Description: A concise example that loops through a list of Excel files, populates column A with numbers 1‑10, assigns a shared formula (e.g., =A1^2) to column B, recalculates all formulas, and overwrites each workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# batch update | shared formula Excel .NET | process multiple workbooks | set shared formula programmatically | recalculate and save Excel files | Aspose.Cells SetSharedFormula example
// Common Searches: batch apply shared formula Aspose.Cells C# | loop through Excel files set formula .NET | how to use SetSharedFormula in multiple workbooks | calculate formulas and save workbooks with Aspose.Cells | C# script to update many Excel files at once
// Developer Intent: Automate the insertion of the same shared formula across a range in each workbook of a collection and persist the changes.
// Use Cases: Populate a numeric series in column A of several spreadsheets and compute its square in column B via a shared formula. | Efficiently propagate identical calculations across rows without writing the formula to every cell. | Recalculate workbook formulas after modification and overwrite the original files in a single batch operation.
// AI Prompts: Write C# code that loads a list of Excel files with Aspose.Cells, fills column A with 1‑10, sets a shared formula in column B, recalculates, and saves each file. | Show how to use Aspose.Cells SetSharedFormula in a loop to batch‑process multiple workbooks. | Explain step‑by‑step how to add data, apply a shared formula, recalculate, and overwrite Excel files using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace BatchSharedFormulaDemo
{
    // A concise example that loops through a list of Excel files, populates column A with numbers 1‑10, assigns a shared formula (e.g., =A1^2) to column B, recalculates all formulas, and overwrites each workbook with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // List of workbook file paths to process
            string[] workbookFiles = new string[]
            {
                "Book1.xlsx",
                "Book2.xlsx",
                "Book3.xlsx"
            };

            // Loop through each workbook
            foreach (string filePath in workbookFiles)
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Example data: populate column A with numbers 1..10
                for (int i = 0; i < 10; i++)
                {
                    cells[i, 0].PutValue(i + 1); // A1..A10
                }

                // Set a shared formula in column B starting from B1
                // Formula: square of the value in column A (e.g., =A1^2)
                // This will propagate to B1:B10
                cells["B1"].SetSharedFormula("=A1^2", 10, 1);

                // Calculate formulas so that values are updated
                workbook.CalculateFormula();

                // Save the workbook (overwrites the original file)
                workbook.Save(filePath);
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
