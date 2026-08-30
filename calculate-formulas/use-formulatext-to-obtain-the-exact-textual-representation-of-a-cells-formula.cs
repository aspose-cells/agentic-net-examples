// Title: Retrieve the exact A1‑style formula text from a worksheet cell with Aspose.Cells GetFormula in C#
// AI Prompts: Generate C# code that uses Aspose.Cells to read a cell's formula as a plain A1‑style string. | Show how to enable Worksheet.ShowFormulas in Aspose.Cells so the worksheet displays formulas instead of calculated values. | Provide a complete example that extracts a formula, prints it to the console, and then saves the workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# GetFormula non localized A1 style example | How to read raw Excel formula string from a cell using Aspose.Cells | Display formulas in an Excel file with Aspose.Cells ShowFormulas property | Save workbook after extracting formula text with Aspose.Cells .NET | GetFormula(false, false) usage in Aspose.Cells C# tutorial
// Tags: retrieve cell formula string Aspose.Cells C# | non‑localized GetFormula A1 style | show formulas worksheet Aspose.Cells | save workbook after reading formula Aspose.Cells | extract raw Excel formula Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, assigns a SUM formula to cell C1, uses GetFormula(false, false) to obtain the exact non‑localized A1‑style formula string, prints it, enables ShowFormulas to display the formula directly in the worksheet, and finally saves the workbook to disk.
    public class FormulaTextDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Set a sample formula in cell C1
                Cell targetCell = cells["C1"];
                targetCell.Formula = "=SUM(A1:B1)";

                // Optionally put values in referenced cells
                cells["A1"].PutValue(5);
                cells["B1"].PutValue(10);

                // Obtain the exact textual representation of the formula
                // GetFormula(false, false) returns the formula in A1 style, non‑localized
                string formulaText = targetCell.GetFormula(false, false);

                // Display the retrieved formula text
                Console.WriteLine("Exact formula text in C1: " + formulaText);

                // Demonstrate that Worksheet.ShowFormulas can display the formula directly in the UI
                worksheet.ShowFormulas = true;
                Console.WriteLine("Cell C1 displayed as formula: " + cells["C1"].StringValue);

                // Save the workbook (lifecycle: save)
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormulaTextDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            FormulaTextDemo.Run();
        }
    }
}
