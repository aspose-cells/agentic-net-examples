// Title: Extract formulas from a named range in an Excel workbook and save them to a text file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells extract formula text from a named range and export to txt | How to write Excel cell formulas to a .txt file with Aspose.Cells .NET | GetRangeByName iterate cells and retrieve formulas in C# | Save named range formulas to a plain text file using Aspose.Cells | Aspose.Cells C# write cell formulas or displayed values from a named range to file
// Tags: Aspose.Cells GetRangeByName formula extraction | export named range formulas to text file C# | cell.IsFormula property usage Aspose.Cells | write Excel formulas with StreamWriter .NET | named range processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExtractFormulasFromNamedRange
{
    // The C# program loads an Excel workbook, obtains a named range via GetRangeByName, iterates each cell, and writes the cell name together with its formula (or displayed value if not a formula) to a text file using StreamWriter.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file
            string excelPath = "input.xlsx";

            // Path to the output text file
            string outputPath = "Formulas.txt";

            // Verify that the input file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file '{excelPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Name of the named range to process
                string namedRangeName = "MyRange";

                // Retrieve the range object by name
                Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName(namedRangeName);
                if (range == null)
                {
                    Console.WriteLine($"Named range '{namedRangeName}' not found.");
                    return;
                }

                // Write formulas/values to the output file
                using (StreamWriter writer = new StreamWriter(outputPath, false))
                {
                    foreach (Cell cell in range)
                    {
                        // Use formula if present; otherwise, use the cell's displayed value
                        string output = cell.IsFormula ? cell.Formula : cell.StringValue;
                        writer.WriteLine($"{cell.Name}: {output}");
                    }
                }

                Console.WriteLine($"Formulas extracted to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
