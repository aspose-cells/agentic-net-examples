// Title: Export Excel formula evaluation order to a text file using Aspose.Cells for .NET
// Common Searches: how to list all formula cells in an Excel workbook with Aspose.Cells C# | export Aspose.Cells formula evaluation sequence to txt for debugging | C# write formula evaluation order of workbook to a text file using Aspose.Cells | retrieve formula dependency order from Excel file with Aspose.Cells .NET
// Tags: Aspose.Cells export formula order to text | enumerate formula cells Aspose.Cells .NET | write Excel formula list to txt C# | debug formula dependency chain Aspose.Cells | iterate worksheets cells Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook with Aspose.Cells, iterates through each worksheet and cell, detects cells that contain formulas, and writes the worksheet name and cell reference to a text file (FormulaEvaluationOrder.txt) to help debug complex formula dependency chains.
class ExportFormulaEvaluationOrder
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Output file path
            string outputPath = "FormulaEvaluationOrder.txt";

            // Write the evaluation order (cells containing formulas) to the text file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                writer.WriteLine("Formula Evaluation Order (cells containing formulas):");
                writer.WriteLine("------------------------------------------------------");

                int index = 1;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all cells in the worksheet
                    foreach (Cell cell in sheet.Cells)
                    {
                        // Check if the cell contains a formula
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            string line = $"{index}: Worksheet=\"{sheet.Name}\", Cell=\"{cell.Name}\"";
                            writer.WriteLine(line);
                            index++;
                        }
                    }
                }
            }

            Console.WriteLine($"Formula evaluation order has been exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
